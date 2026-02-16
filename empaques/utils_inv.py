# empaques/utils_inv.py
from decimal import Decimal
from django.db import transaction
from .models import InventoryItem, InventoryMovement, Presentation

# empaques/utils_inv.py
from decimal import Decimal
from django.db import transaction
from .models import InventoryItem, InventoryMovement, Presentation

def ensure_item_for_presentation(pres: Presentation) -> InventoryItem:
    # SKU distinto solo para arándano
    sku = f"AR-CS-{pres.id:04d}"
    item, _ = InventoryItem.objects.get_or_create(
        sku=sku,
        defaults={"name": f"Clamshells {pres.name}", "unit": "cs"}
    )
    return item

@transaction.atomic
def post_out_for_shipment(pres: Presentation, boxes: int, date, shipment_ref="", user=None, notes="Salida por embarque"):
    # ← BLOQUEADOR: si NO es arándano, no toques el almacén de arándano
    if not getattr(pres, "is_arandano", False):
        return

    cs_por_caja = int(getattr(pres, "cs_por_caja", 0))
    cs = cs_por_caja * int(boxes or 0)
    if cs <= 0:
        return

    item = ensure_item_for_presentation(pres)  # solo arándano
    InventoryMovement.objects.create(
        item=item,
        date=date,
        type="OUT",
        quantity=Decimal(cs),
        reference=shipment_ref,
        notes=notes,
        created_by=user
    )

def cs_per_box(pres: Presentation) -> int:
    """
    Devuelve cuántos clamshells por caja tiene esta presentación.
    Si usas 6/9.8/18oz puras, puedes guardar esto en Presentation o en otro lado.
    """
    # Si Presentation tiene campos por caja (recomendado):
    #   pres.cs_6oz_por_caja, pres.cs_9_8oz_por_caja, pres.cs_18oz_por_caja
    # y cada Presentation representa SOLO una de esas variantes, guarda en un único campo pres.cs_por_caja.
    # Para este ejemplo, supongamos que Presentation tiene un campo genérico:
    return int(getattr(pres, "cs_por_caja", 0))  # AJUSTA según tu modelo

@transaction.atomic
def post_in_from_field(pres: Presentation, cs_amount: int, date, user=None, reference="Entrada desde Campo", notes=""):
    """
    Entrada al almacén de Empaque (en cs) cuando Campo manda clamshells a Empaque.
    """
    if cs_amount <= 0:
        return
    item = ensure_item_for_presentation(pres)
    InventoryMovement.objects.create(
        item=item,
        date=date,
        type="IN",
        quantity=Decimal(cs_amount),
        reference=reference,
        notes=notes,
        created_by=user
    )

@transaction.atomic
def post_out_for_shipment(pres: Presentation, boxes: int, date, shipment_ref="", user=None, notes="Salida por embarque"):
    """
    Salida del almacén de Empaque (en cs), a partir de cajas de una Presentation.
    """
    cs = cs_per_box(pres) * int(boxes or 0)
    if cs <= 0:
        return
    item = ensure_item_for_presentation(pres)
    InventoryMovement.objects.create(
        item=item,
        date=date,
        type="OUT",
        quantity=Decimal(cs),
        reference=shipment_ref,
        notes=notes,
        created_by=user
    )