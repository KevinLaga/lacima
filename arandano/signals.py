# arandano/signals.py
from decimal import Decimal
from django.db.models.signals import post_save, post_delete
from django.dispatch import receiver
from django.contrib.contenttypes.models import ContentType

from .models import (
    ProduccionDia, ProduccionItem,
    SalidaDia, SalidaItem, DestinoSalida,
    InventarioMovimiento, ScopeInv, ConceptoInv
)

def _ct_for(obj):
    return ContentType.objects.get_for_model(obj.__class__)

# ---- Helpers de upsert para evitar duplicados si se edita un documento ----

def _delete_movs_for(instance):
    ct = _ct_for(instance)
    InventarioMovimiento.objects.filter(content_type=ct, object_id=instance.pk).delete()

# ========== Producción (Campo) ==========
@receiver(post_save, sender=ProduccionDia)
def producciondia_to_mov(sender, instance: ProduccionDia, created, **kwargs):
    # Borramos y re-creamos (más simple y consistente al editar)
    _delete_movs_for(instance)

    items = ProduccionItem.objects.filter(produccion=instance).select_related("variedad")
    for it in items:
        if not it.variedad:
            continue
        InventarioMovimiento.objects.create(
            scope=ScopeInv.CAMPO,
            campo=instance.campo,
            variedad=it.variedad,
            fecha=instance.fecha,
            concepto=ConceptoInv.ENTRADA_CAMPO,
            kg=Decimal(str(it.kg or 0)),
            cs_6oz=int(it.cs_6oz or 0),
            cs_9_8oz=int(it.cs_9_8oz or 0),
            cs_18oz=int(it.cs_18oz or 0),
            notas=instance.notas or "",
            content_type=_ct_for(instance),
            object_id=instance.pk,
        )

@receiver(post_delete, sender=ProduccionDia)
def producciondia_del(sender, instance, **kwargs):
    _delete_movs_for(instance)

# ========== Salida (Campo → Otro / Campo → Empaque) ==========
@receiver(post_save, sender=SalidaDia)
def salidadia_to_mov(sender, instance: SalidaDia, created, **kwargs):
    _delete_movs_for(instance)

    items = SalidaItem.objects.filter(salida=instance).select_related("variedad")
    for it in items:
        if not it.variedad:
            continue

        kg   = it.kg or 0
        cs6  = int(it.cs_6oz   or 0)
        cs98 = int(it.cs_9_8oz or 0)
        cs18 = int(it.cs_18oz  or 0)

        if instance.destino == DestinoSalida.OTRO:
            # Solo baja de Campo
            InventarioMovimiento.objects.create(
                scope=ScopeInv.CAMPO,
                campo=instance.campo,
                variedad=it.variedad,
                fecha=instance.fecha,
                concepto=ConceptoInv.SALIDA_CAMPO_OTRO,
                kg=kg, cs_6oz=cs6, cs_9_8oz=cs98, cs_18oz=cs18,
                notas=(instance.destino_detalle or instance.notas or ""),
                content_type=_ct_for(instance),
                object_id=instance.pk,
            )
        else:
            # Transferencia Campo→Empaque: dos movimientos espejo
            # 1) Baja Campo
            InventarioMovimiento.objects.create(
                scope=ScopeInv.CAMPO,
                campo=instance.campo,
                variedad=it.variedad,
                fecha=instance.fecha,
                concepto=ConceptoInv.TRANSFER_CAMPO_A_EMPAQUE,
                kg=kg, cs_6oz=cs6, cs_9_8oz=cs98, cs_18oz=cs18,
                notas="Transferencia a Empaque",
                content_type=_ct_for(instance),
                object_id=instance.pk,
            )
            # 2) Sube Empaque (campo NULL porque es inventario global de empaque)
            InventarioMovimiento.objects.create(
                scope=ScopeInv.EMPAQUE,
                campo=None,
                variedad=it.variedad,
                fecha=instance.fecha,
                concepto=ConceptoInv.ENTRADA_EMPAQUE_DESDE_CAMPO,
                kg=kg, cs_6oz=cs6, cs_9_8oz=cs98, cs_18oz=cs18,
                notas=f"Desde {instance.campo}",
                content_type=_ct_for(instance),
                object_id=instance.pk,
            )

@receiver(post_delete, sender=SalidaDia)
def salidadia_del(sender, instance, **kwargs):
    _delete_movs_for(instance)