# empaques/views_inventory.py
from decimal import Decimal

from django.contrib.auth.decorators import login_required, permission_required
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.http import HttpResponseForbidden

from django.db.models import Sum, Q, F, DecimalField, Value as V, ExpressionWrapper
from django.db.models.functions import Coalesce

from .models import InventoryItem, InventoryMovement
from .forms_inventory import InventoryItemForm, InventoryMovementForm


@login_required
@permission_required('empaques.view_inventoryitem', raise_exception=True)
def almacen_list(request):
    """
    Lista de artículos con stock calculado (ent - sal + adj) vía annotate.
    """
    q = (request.GET.get("q") or "").strip()

    items_qs = InventoryItem.objects.all()
    if q:
        items_qs = items_qs.filter(
            Q(sku__icontains=q) | Q(name__icontains=q) | Q(location__icontains=q)
        )

    items = (
        items_qs
        .annotate(
            ent=Coalesce(
                Sum(
                    'movements__quantity',
                    filter=Q(movements__type='IN'),
                    output_field=DecimalField(max_digits=12, decimal_places=2),
                ),
                V(0),
                output_field=DecimalField(max_digits=12, decimal_places=2),
            ),
            sal=Coalesce(
                Sum(
                    'movements__quantity',
                    filter=Q(movements__type='OUT'),
                    output_field=DecimalField(max_digits=12, decimal_places=2),
                ),
                V(0),
                output_field=DecimalField(max_digits=12, decimal_places=2),
            ),
            adj=Coalesce(
                Sum(
                    'movements__quantity',
                    filter=Q(movements__type='ADJ'),
                    output_field=DecimalField(max_digits=12, decimal_places=2),
                ),
                V(0),
                output_field=DecimalField(max_digits=12, decimal_places=2),
            ),
        )
        .annotate(
            stock_qty=ExpressionWrapper(
                F('ent') - F('sal') + F('adj'),
                output_field=DecimalField(max_digits=12, decimal_places=2),
            )
        )
        .order_by('name')
    )

    # Alta rápida (opcional)
    if request.method == 'POST' and 'add_item' in request.POST:
        form = InventoryItemForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect(reverse('almacen_list'))
    else:
        form = InventoryItemForm()

    return render(request, "empaques/almacen_list.html", {
        "items": items,
        "form": form,
        "q": q,
    })


from django.contrib import messages

# views_inventory.py
from django.contrib import messages
from django.shortcuts import render, redirect
from django.urls import reverse

@login_required
@permission_required("empaques.add_inventoryitem", raise_exception=True)
def almacen_item_new(request):
    if request.method == "POST":
        form = InventoryItemForm(request.POST)
        if form.is_valid():
            item = form.save(commit=False)

            # Rellena created_by si existe ese campo en el modelo
            if hasattr(item, "created_by_id") and not item.created_by_id:
                item.created_by = request.user

            item.save()
            messages.success(request, "Artículo creado correctamente.")
            return redirect(reverse("almacen_list"))
        else:
            # Para que veas rápidamente si llega aquí
            messages.error(request, "Revisa los errores del formulario.")
    else:
        form = InventoryItemForm()

    return render(request, "empaques/almacen_item_form.html", {"form": form})


from django.contrib import messages

@login_required
@permission_required("empaques.add_inventorymovement", raise_exception=True)
def almacen_movement_new(request, tipo=None):
    """
    Crea un movimiento. Si es ADJ, la cantidad capturada se interpreta como
    'existencia final deseada' y se convierte a delta (= deseada - stock actual).
    """
    if request.method == "POST":
        form = InventoryMovementForm(request.POST, user=request.user)
        if form.is_valid():
            m = form.save(commit=False)

            # Si es ajuste: convertir cantidad a DELTA para que el stock final sea la cantidad deseada
            if m.type == "ADJ":
                agg = InventoryMovement.objects.filter(item=m.item).aggregate(
                    ent=Coalesce(Sum('quantity', filter=Q(type='IN'),
                                     output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                                 output_field=DecimalField(max_digits=12, decimal_places=2)),
                    sal=Coalesce(Sum('quantity', filter=Q(type='OUT'),
                                     output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                                 output_field=DecimalField(max_digits=12, decimal_places=2)),
                    adj=Coalesce(Sum('quantity', filter=Q(type='ADJ'),
                                     output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                                 output_field=DecimalField(max_digits=12, decimal_places=2)),
                )
                current_stock = (agg["ent"] - agg["sal"] + agg["adj"]) or Decimal("0")
                desired_final = m.quantity  # lo que el usuario escribió en el campo
                delta = desired_final - current_stock
                m.quantity = delta
                if not m.notes:
                    m.notes = f"Ajuste a {desired_final} (Δ {delta})"

            m.save()
            messages.success(request, "Movimiento registrado correctamente.")
            return redirect("almacen_kardex", item_id=m.item_id)
        else:
            messages.error(request, "Revisa los errores del formulario.")
    else:
        initial = {}
        if tipo in ("IN", "OUT", "ADJ"):
            initial["type"] = tipo
        # ← PRESELECCIONA el artículo si venimos de la lista con ?item=ID
        item_id = request.GET.get("item")
        if item_id:
            initial["item"] = item_id

        form = InventoryMovementForm(initial=initial)

    return render(request, "empaques/almacen_movement_form.html", {"form": form})

@login_required
def almacen_kardex(request, item_id):
    """
    Historial de movimientos del artículo + saldo corrido.
    """
    if not request.user.has_perm("empaques.view_inventorymovement"):
        return HttpResponseForbidden("No tienes permiso para ver movimientos.")

    item = get_object_or_404(InventoryItem, pk=item_id)

    # Movimientos en orden cronológico
    movimientos = (
        item.movements
        .select_related("created_by")
        .order_by("date", "id")
    )

    # Saldo corrido
    saldo = Decimal("0")
    rows = []
    for m in movimientos:
        if m.type == "IN":
            delta = m.quantity
        elif m.type == "OUT":
            delta = -m.quantity
        else:  # ADJ
            delta = m.quantity
        saldo += delta
        rows.append({
            "m": m,
            "saldo": saldo
        })

    # Stock actual (puedes usar item.stock también)
    stock_agg = item.movements.aggregate(
        ent=Coalesce(Sum('quantity', filter=Q(type='IN'),  output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                     output_field=DecimalField(max_digits=12, decimal_places=2)),
        sal=Coalesce(Sum('quantity', filter=Q(type='OUT'), output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                     output_field=DecimalField(max_digits=12, decimal_places=2)),
        adj=Coalesce(Sum('quantity', filter=Q(type='ADJ'), output_field=DecimalField(max_digits=12, decimal_places=2)), V(0),
                     output_field=DecimalField(max_digits=12, decimal_places=2)),
    )
    stock_qty = stock_agg["ent"] - stock_agg["sal"] + stock_agg["adj"]

    return render(request, "empaques/almacen_kardex.html", {
        "item": item,
        "rows": rows,         # ← movimientos + saldo
        "stock_qty": stock_qty,
    })
