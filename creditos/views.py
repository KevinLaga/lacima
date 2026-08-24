from collections import OrderedDict
from datetime import timedelta
from decimal import Decimal

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Prefetch
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_http_methods

from .forms import AbonoForm, CreditoForm, GarantiaForm
from .models import (
    BANCO_CHOICES, EMPRESA_CHOICES, MONEDA_CHOICES, SIMBOLO_MONEDA,
    TIPO_CREDITO_CHOICES, Abono, Credito, Garantia,
)

# Menú de proximidad a vencer. clave -> (etiqueta, días); None = sin tope.
VENCE_FILTROS = OrderedDict([
    ("todos",    ("Todos",            None)),
    ("vencidos", ("Vencidos",         -1)),
    ("30",       ("Vencen en 30 días", 30)),
    ("60",       ("Vencen en 60 días", 60)),
    ("90",       ("Vencen en 90 días", 90)),
])

ORDEN_OPCIONES = OrderedDict([
    ("vencimiento", "Más próximos a vencer"),
    ("saldo",       "Mayor saldo"),
    ("monto",       "Mayor monto"),
    ("reciente",    "Disposición más reciente"),
])


def _totales_por_moneda(creditos):
    """Suma monto/abonado/saldo agrupado por moneda (no se mezclan divisas)."""
    acc = {}
    for c in creditos:
        d = acc.setdefault(c.moneda, {
            "moneda": c.moneda,
            "simbolo": SIMBOLO_MONEDA.get(c.moneda, "$"),
            "monto": Decimal("0"),
            "abonado": Decimal("0"),
            "saldo": Decimal("0"),
            "n": 0,
        })
        d["monto"]   += Decimal(c.monto or 0)
        d["abonado"] += c.total_abonado
        d["saldo"]   += c.saldo
        d["n"]       += 1
    return [acc[k] for k in sorted(acc)]


@login_required
def credito_list(request):
    hoy = timezone.localdate()

    empresa   = (request.GET.get("empresa") or "").strip()
    banco     = (request.GET.get("banco") or "").strip()
    moneda    = (request.GET.get("moneda") or "").strip()
    vence     = (request.GET.get("vence") or "todos").strip()
    orden     = (request.GET.get("orden") or "vencimiento").strip()
    ocultar_liq = request.GET.get("ocultar_liquidados") == "1"

    if vence not in VENCE_FILTROS:
        vence = "todos"
    if orden not in ORDEN_OPCIONES:
        orden = "vencimiento"

    qs = Credito.objects.select_related("garantia").prefetch_related(
        Prefetch("abonos", queryset=Abono.objects.order_by("-fecha", "-id"))
    )

    if empresa:
        qs = qs.filter(empresa=empresa)
    if banco:
        qs = qs.filter(banco=banco)
    if moneda:
        qs = qs.filter(moneda=moneda)

    # Filtro por proximidad de vencimiento (a nivel de base de datos)
    dias = VENCE_FILTROS[vence][1]
    if vence == "vencidos":
        qs = qs.filter(fecha_vencimiento__lt=hoy)
    elif dias is not None:
        qs = qs.filter(fecha_vencimiento__gte=hoy,
                       fecha_vencimiento__lte=hoy + timedelta(days=dias))

    creditos = list(qs)

    # 'liquidado' y 'saldo' dependen de los abonos: se resuelven en Python
    if ocultar_liq:
        creditos = [c for c in creditos if not c.liquidado]

    if orden == "saldo":
        creditos.sort(key=lambda c: c.saldo, reverse=True)
    elif orden == "monto":
        creditos.sort(key=lambda c: c.monto or 0, reverse=True)
    elif orden == "reciente":
        creditos.sort(key=lambda c: c.fecha_disposicion or hoy, reverse=True)
    else:
        # Más próximos a vencer primero; los liquidados al final
        creditos.sort(key=lambda c: (c.liquidado, c.fecha_vencimiento or hoy))

    # Conteos para las pestañas del menú
    todos = list(Credito.objects.all())
    pendientes = [c for c in todos if not c.liquidado]
    conteos = {
        "todos":    len(todos),
        "vencidos": sum(1 for c in pendientes if (c.dias_para_vencer or 0) < 0),
        "30":       sum(1 for c in pendientes if 0 <= (c.dias_para_vencer or -1) <= 30),
        "60":       sum(1 for c in pendientes if 0 <= (c.dias_para_vencer or -1) <= 60),
        "90":       sum(1 for c in pendientes if 0 <= (c.dias_para_vencer or -1) <= 90),
    }

    return render(request, "creditos/credito_list.html", {
        "creditos": creditos,
        "totales": _totales_por_moneda(creditos),
        "hoy": hoy,
        "f": {
            "empresa": empresa, "banco": banco, "moneda": moneda,
            "vence": vence, "orden": orden,
            "ocultar_liquidados": ocultar_liq,
        },
        "empresas": EMPRESA_CHOICES,
        "bancos": BANCO_CHOICES,
        "monedas": MONEDA_CHOICES,
        "vence_filtros": [(k, v[0], conteos.get(k, 0)) for k, v in VENCE_FILTROS.items()],
        "orden_opciones": list(ORDEN_OPCIONES.items()),
    })


@login_required
@require_http_methods(["GET", "POST"])
def credito_create(request):
    if request.method == "POST":
        form = CreditoForm(request.POST)
        if form.is_valid():
            credito = form.save()
            messages.success(request, "Crédito registrado correctamente.")
            return redirect("creditos:credito_detail", pk=credito.pk)
        messages.error(request, "Revisa los campos marcados.")
    else:
        form = CreditoForm()

    return render(request, "creditos/credito_form.html", {
        "form": form, "credito": None, "titulo": "Nuevo crédito",
    })


@login_required
@require_http_methods(["GET", "POST"])
def credito_edit(request, pk):
    credito = get_object_or_404(Credito, pk=pk)

    if request.method == "POST":
        form = CreditoForm(request.POST, instance=credito)
        if form.is_valid():
            form.save()
            messages.success(request, "Crédito actualizado.")
            return redirect("creditos:credito_detail", pk=credito.pk)
        messages.error(request, "Revisa los campos marcados.")
    else:
        form = CreditoForm(instance=credito)

    return render(request, "creditos/credito_form.html", {
        "form": form, "credito": credito, "titulo": "Editar crédito",
    })


@login_required
@require_http_methods(["GET", "POST"])
def credito_detail(request, pk):
    credito = get_object_or_404(
        Credito.objects.select_related("garantia"), pk=pk,
    )

    if request.method == "POST":
        abono_form = AbonoForm(request.POST, credito=credito)
        if abono_form.is_valid():
            abono = abono_form.save(commit=False)
            abono.credito = credito
            abono.save()
            messages.success(request, f"Abono de {abono.monto_fmt} registrado.")
            return redirect("creditos:credito_detail", pk=credito.pk)
        messages.error(request, "No se pudo registrar el abono.")
    else:
        abono_form = AbonoForm(credito=credito)

    return render(request, "creditos/credito_detail.html", {
        "credito": credito,
        "abonos": credito.abonos.all(),
        "abono_form": abono_form,
    })


@login_required
@require_http_methods(["POST"])
def abono_delete(request, pk):
    abono = get_object_or_404(Abono, pk=pk)
    credito_pk = abono.credito_id
    abono.delete()
    messages.success(request, "Abono eliminado.")
    return redirect("creditos:credito_detail", pk=credito_pk)


@login_required
@require_http_methods(["GET", "POST"])
def garantia_list(request):
    """Catálogo de campos/terrenos dados en garantía."""
    if request.method == "POST":
        form = GarantiaForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, "Garantía agregada.")
            return redirect("creditos:garantia_list")
        messages.error(request, "Revisa los campos marcados.")
    else:
        form = GarantiaForm()

    return render(request, "creditos/garantia_list.html", {
        "form": form,
        "garantias": Garantia.objects.all().order_by("nombre"),
    })
