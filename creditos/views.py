from collections import OrderedDict
from datetime import timedelta
from decimal import Decimal

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Prefetch
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.utils.text import slugify
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


def _aplicar_filtros(request):
    """
    Resuelve los filtros de la lista de créditos.

    La usan tanto la pantalla como la exportación a Excel, para que el archivo
    descargado sea exactamente lo que el usuario está viendo.

    Devuelve (creditos_ordenados, dict_de_filtros).
    """
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

    filtros = {
        "empresa": empresa, "banco": banco, "moneda": moneda,
        "vence": vence, "orden": orden, "ocultar_liquidados": ocultar_liq,
    }
    return creditos, filtros


@login_required
def credito_list(request):
    hoy = timezone.localdate()
    creditos, filtros = _aplicar_filtros(request)

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

    # Querystring actual, para que el botón de Excel exporte lo mismo que se ve
    qs_actual = request.GET.urlencode()

    return render(request, "creditos/credito_list.html", {
        "creditos": creditos,
        "totales": _totales_por_moneda(creditos),
        "hoy": hoy,
        "f": filtros,
        "qs_actual": qs_actual,
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
def credito_export_xlsx(request):
    """
    Exporta a Excel EXACTAMENTE los créditos que se están viendo en la lista:
    mismos filtros, mismo orden y mismas columnas.
    """
    from io import BytesIO

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    creditos, filtros = _aplicar_filtros(request)
    hoy = timezone.localdate()

    wb = Workbook()
    ws = wb.active
    ws.title = "Créditos"

    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    th_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    th_fill = PatternFill("solid", fgColor="1E3A5F")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    # ── Título y filtros aplicados ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    c = ws.cell(row=1, column=1, value="Créditos")
    c.font = Font(name="Calibri", size=16, bold=True, color="1E3A5F")
    c.alignment = Alignment(horizontal="left", vertical="center")

    partes = [f"Generado {hoy.strftime('%d/%m/%Y')}"]
    if filtros["empresa"]:
        partes.append("Empresa: " + dict(EMPRESA_CHOICES).get(filtros["empresa"], filtros["empresa"]))
    if filtros["banco"]:
        partes.append("Banco: " + dict(BANCO_CHOICES).get(filtros["banco"], filtros["banco"]))
    if filtros["moneda"]:
        partes.append("Moneda: " + filtros["moneda"])
    partes.append("Vencimiento: " + VENCE_FILTROS[filtros["vence"]][0])
    partes.append("Orden: " + ORDEN_OPCIONES[filtros["orden"]])
    if filtros["ocultar_liquidados"]:
        partes.append("Sin liquidados")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=13)
    c = ws.cell(row=2, column=1, value="  ·  ".join(partes))
    c.font = Font(name="Calibri", size=10, italic=True, color="6D6D6D")
    c.alignment = Alignment(horizontal="left", vertical="center")

    # ── Encabezados (mismas columnas que la tabla en pantalla) ──
    headers = [
        ("Empresa", 20), ("Banco", 20), ("Tipo", 24), ("Garantía", 22),
        ("Moneda", 9), ("Monto", 16), ("Abonado", 16), ("Saldo", 16),
        ("Tasa (%)", 11), ("Plazo (meses)", 13),
        ("Disposición", 13), ("Vencimiento", 13), ("Estado", 13),
    ]
    r = 4
    for col, (h, w) in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=col, value=h)
        cell.font = th_font
        cell.fill = th_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[r].height = 24
    r += 1

    fmt_money = '#,##0.00'
    fill_vencido = PatternFill("solid", fgColor="FDECEC")

    for cr in creditos:
        vals = [
            cr.get_empresa_display(),
            cr.get_banco_display(),
            cr.tipo_credito_label,
            cr.garantia.nombre if cr.garantia else "",
            cr.moneda,
            float(cr.monto or 0),
            float(cr.total_abonado),
            float(cr.saldo),
            float(cr.tasa) if cr.tasa is not None else None,
            cr.plazo_meses,
            cr.fecha_disposicion,
            cr.fecha_vencimiento,
            cr.estado_label,
        ]
        for col, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=col, value=v)
            cell.border = border
            if col in (6, 7, 8):
                cell.number_format = fmt_money
                cell.alignment = right
            elif col == 9:
                cell.number_format = '0.000'
                cell.alignment = right
            elif col in (10, 13):
                cell.alignment = center
            elif col in (11, 12):
                cell.number_format = 'dd/mm/yyyy'
                cell.alignment = center
            if cr.estado == "vencido":
                cell.fill = fill_vencido
        r += 1

    if not creditos:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        c = ws.cell(row=r, column=1, value="No hay créditos que coincidan con el filtro.")
        c.alignment = center
        c.font = Font(italic=True, color="9CA3AF")
        r += 1

    # ── Totales por moneda (no se mezclan divisas) ──
    r += 1
    for t in _totales_por_moneda(creditos):
        ws.cell(row=r, column=1, value=f"TOTAL {t['moneda']}").font = Font(bold=True)
        ws.cell(row=r, column=5, value=f"{t['n']} créditos").alignment = center
        for col, key in ((6, "monto"), (7, "abonado"), (8, "saldo")):
            cell = ws.cell(row=r, column=col, value=float(t[key]))
            cell.font = Font(bold=True)
            cell.number_format = fmt_money
            cell.alignment = right
            cell.border = border
        r += 1

    ws.freeze_panes = "A5"

    out = BytesIO()
    wb.save(out)
    out.seek(0)

    partes_nombre = ["creditos"]
    if filtros["empresa"]:
        partes_nombre.append(slugify(filtros["empresa"]))
    if filtros["banco"]:
        partes_nombre.append(slugify(filtros["banco"]))
    if filtros["vence"] != "todos":
        partes_nombre.append(filtros["vence"])
    partes_nombre.append(hoy.isoformat())
    fname = "_".join(partes_nombre) + ".xlsx"

    resp = HttpResponse(
        out.read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="{fname}"'
    return resp


@login_required
@require_http_methods(["POST"])
def garantia_delete(request, pk):
    garantia = get_object_or_404(Garantia, pk=pk)

    n = garantia.creditos.count()
    if n:
        messages.error(
            request,
            f"No se puede quitar «{garantia.nombre}»: está usada por {n} "
            f"crédito{'s' if n != 1 else ''}."
        )
    else:
        nombre = garantia.nombre
        garantia.delete()
        messages.success(request, f"Garantía «{nombre}» eliminada.")

    return redirect("creditos:garantia_list")


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
