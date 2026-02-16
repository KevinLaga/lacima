import os
from django.contrib.auth import get_user_model
from django.db.models.signals import post_migrate
# embarques/signals.py
from django.db.models.signals import post_save

from .models import DetalleEmbarque
from .utils_conv import clamshells_y_kg_de_presentacion
from decimal import Decimal

from django.db.models.signals import pre_save, post_save, post_delete
from django.dispatch import receiver

from .models import ShipmentItem, InventoryItem, InventoryMovement


# Mapeo: Presentation.name -> SKU de almacén y cuántos clamshells por caja
ARANDANO_MAP = {
    "4.5 LBS (12X6oz)":   {"sku": "AR-CS6",  "cs_per_box": 12},
    "7.35 LBS (12X9.8oz)": {"sku": "AR-CS98", "cs_per_box": 12},
    "13.5 LBS (12X18oz)": {"sku": "AR-CS18", "cs_per_box": 12},
}


def _get_arandano_cfg(instance: ShipmentItem):
    pres = instance.presentation
    if not pres:
        return None
    return ARANDANO_MAP.get((pres.name or "").strip())


def _movement_ref(item: ShipmentItem) -> str:
    ship = item.shipment
    tn = getattr(ship, "tracking_number", "") or ""
    return f"Embarque {tn}".strip() if tn else f"Embarque #{ship.pk}".strip()


@receiver(pre_save, sender=ShipmentItem)
def shipmentitem_store_old(sender, instance: ShipmentItem, **kwargs):
    """
    Guardamos valores anteriores para calcular delta cuando editan quantity/presentation.
    """
    instance._old_qty = 0
    instance._old_pres_name = ""
    if instance.pk:
        old = ShipmentItem.objects.filter(pk=instance.pk).select_related("presentation").first()
        if old:
            instance._old_qty = int(old.quantity or 0)
            instance._old_pres_name = (old.presentation.name or "").strip() if old.presentation else ""


@receiver(post_save, sender=ShipmentItem)
def shipmentitem_post_to_arandano_inventory(sender, instance: ShipmentItem, created, **kwargs):
    """
    Si el ShipmentItem es de una presentación de arándano:
    - Cuando se crea: OUT (cajas * cs_per_box)
    - Cuando se edita: OUT/IN por el delta (si aumentó => OUT, si bajó => IN)
    """
    cfg_new = _get_arandano_cfg(instance)

    # Si el item NO es de arándano, no hacemos nada
    old_name = getattr(instance, "_old_pres_name", "") or ""
    cfg_old = ARANDANO_MAP.get(old_name)

    # Si no era arándano antes y tampoco ahora -> nada
    if not cfg_old and not cfg_new:
        return

    ship = instance.shipment
    mov_date = ship.date
    user = getattr(ship, "created_by", None)  # si existe, si no, queda None

    # Calcula clamshells nuevos/anteriores
    new_qty = int(instance.quantity or 0)
    old_qty = int(getattr(instance, "_old_qty", 0) or 0)

    # Si cambió de presentación (ej. arándano->no arándano o viceversa), revertimos lo anterior y aplicamos lo nuevo
    if cfg_old and (not cfg_new or cfg_old["sku"] != cfg_new["sku"]):
        # Revertir lo anterior (meter stock de vuelta) = IN
        cs_old = old_qty * int(cfg_old["cs_per_box"])
        if cs_old > 0:
            inv_item, _ = InventoryItem.objects.get_or_create(
                sku=cfg_old["sku"],
                defaults={
                    "name": f"Arándano Clamshell {cfg_new['sku']}",
                    "unit": "cs",
                    "location": "Empaque",
                    "min_stock": Decimal("0"),
                }
            )
            InventoryMovement.objects.create(
                item=inv_item,
                date=mov_date,
                type="IN",
                quantity=Decimal(cs_old),
                reference=f"Reversión por edición (embarque {ship.tracking_number})",
                notes="Se ajustó línea de embarque (presentación/cantidad).",
                created_by=user,
            )

    if cfg_new:
        cs_new = new_qty * int(cfg_new["cs_per_box"])
        cs_old_equiv = 0
        if cfg_old and cfg_old["sku"] == cfg_new["sku"]:
            cs_old_equiv = old_qty * int(cfg_old["cs_per_box"])

        delta = cs_new - cs_old_equiv

        if delta == 0:
            return

        inv_item, _ = InventoryItem.objects.get_or_create(
            sku=cfg_new["sku"],
            defaults={
                "name": f"Arándano Clamshell {cfg_new['sku']}",
                "unit": "cs",
                "location": "Empaque",
                "min_stock": Decimal("0"),
            }
        )
        if delta > 0:
            # Aumentó lo embarcado -> SALE del almacén (OUT)
            InventoryMovement.objects.create(
                item=inv_item,
                date=mov_date,
                type="OUT",
                quantity=Decimal(delta),
                reference=f"Embarque {ship.tracking_number}",
                notes=f"Salida por embarque.",
                created_by=user,
            )
        else:
            # Disminuyó lo embarcado -> regresa al almacén (IN)
            InventoryMovement.objects.create(
                item=inv_item,
                date=mov_date,
                type="IN",
                quantity=Decimal(abs(delta)),
                reference=f"Ajuste Embarque {ship.tracking_number}",
                notes=f"Auto IN por reducción/edición (línea #{instance.pk}).",
                created_by=user,
            )


@receiver(post_delete, sender=ShipmentItem)
def shipmentitem_restore_arandano_inventory(sender, instance: ShipmentItem, **kwargs):
    """
    Si borran una línea de embarque de arándano, regresamos stock (IN).
    """
    cfg = _get_arandano_cfg(instance)
    if not cfg:
        return

    ship = instance.shipment
    mov_date = ship.date
    qty = int(instance.quantity or 0)
    cs = qty * int(cfg["cs_per_box"])
    if cs <= 0:
        return

    inv_item = InventoryItem.objects.get(sku=cfg["sku"])
    InventoryMovement.objects.create(
        item=inv_item,
        date=mov_date,
        type="IN",
        quantity=Decimal(cs),
        reference=f"Eliminación Embarque {ship.tracking_number}",
        notes=f"Auto IN por eliminar línea de embarque (línea #{instance.pk}).",
    )


@receiver(post_migrate)
def create_initial_superuser(sender, **kwargs):
    """
    Crea un superusuario si no existe y hay variables de entorno definidas.
    Se ejecuta después de 'migrate'. Idempotente por username.
    """
    username = os.getenv("DJANGO_SUPERUSER_USERNAME")
    email    = os.getenv("DJANGO_SUPERUSER_EMAIL", "")
    password = os.getenv("DJANGO_SUPERUSER_PASSWORD")

    if not username or not password:
        return  # no hay datos -> no hace nada

    User = get_user_model()
    if not User.objects.filter(username=username).exists():
        User.objects.create_superuser(username=username, email=email, password=password)
