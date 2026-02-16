# arandano/utils_inv.py
from __future__ import annotations
from decimal import Decimal, ROUND_HALF_UP
from django.db.models import Sum
from django.utils.timezone import now
from django.db import models

from .models import Campo, ProduccionItem, SalidaItem

# =========================
# Constantes / helpers
# =========================
Q5 = Decimal("0.00001")

def q5(x) -> Decimal:
    """Redondeo a 5 decimales como Decimal."""
    if x is None:
        x = Decimal("0")
    if not isinstance(x, Decimal):
        x = Decimal(str(x))
    return x.quantize(Q5, rounding=ROUND_HALF_UP)


class ScopeInv(models.TextChoices):
    """Ámbitos para diferenciar inventarios/reportes."""
    CAMPO   = "campo",   "Campo"
    EMPAQUE = "empaque", "Empaque"


# =========================
# Stock por CAMPO/Variedad
# =========================
def stock_por_campo(campo: Campo, hasta_fecha=None):
    """
    Stock por VARIEDAD dentro de un CAMPO:
      stock = (producción hasta fecha) - (salidas hasta fecha)
    Retorna: { variedad_id: {"kg": Decimal, "cs6": int, "cs98": int, "cs18": int} }
    """
    if hasta_fecha is None:
        hasta_fecha = now().date()

    # Entradas (producción)
    prod_rows = (
        ProduccionItem.objects
        .filter(produccion__campo=campo, produccion__fecha__lte=hasta_fecha)
        .values("variedad")
        .annotate(
            kg_total=Sum("kg"),
            cs6_total=Sum("cs_6oz"),
            cs98_total=Sum("cs_9_8oz"),
            cs18_total=Sum("cs_18oz"),
        )
    )
    producido = {
        row["variedad"]: {
            "kg":   q5(row["kg_total"] or 0),
            "cs6":  int(row["cs6_total"] or 0),
            "cs98": int(row["cs98_total"] or 0),
            "cs18": int(row["cs18_total"] or 0),
        }
        for row in prod_rows
    }

    # Salidas
    sal_rows = (
        SalidaItem.objects
        .filter(salida__campo=campo, salida__fecha__lte=hasta_fecha)
        .values("variedad")
        .annotate(
            kg_total=Sum("kg"),
            cs6_total=Sum("cs_6oz"),
            cs98_total=Sum("cs_9_8oz"),
            cs18_total=Sum("cs_18oz"),
        )
    )
    salidas = {
        row["variedad"]: {
            "kg":   q5(row["kg_total"] or 0),
            "cs6":  int(row["cs6_total"] or 0),
            "cs98": int(row["cs98_total"] or 0),
            "cs18": int(row["cs18_total"] or 0),
        }
        for row in sal_rows
    }

    # Stock = producido - salidas
    result = {}
    all_var_ids = set(producido.keys()) | set(salidas.keys())
    for vid in all_var_ids:
        p = producido.get(vid, {"kg": Decimal("0"), "cs6": 0, "cs98": 0, "cs18": 0})
        s = salidas.get(vid,  {"kg": Decimal("0"), "cs6": 0, "cs98": 0, "cs18": 0})
        result[vid] = {
            "kg":   q5(p["kg"]) - q5(s["kg"]),
            "cs6":  int(p["cs6"])  - int(s["cs6"]),
            "cs98": int(p["cs98"]) - int(s["cs98"]),
            "cs18": int(p["cs18"]) - int(s["cs18"]),
        }

    return result


# =========================
# Proxy a post_ledger (empaques)
# =========================
# Si empaques tiene su propio utils_inv.post_ledger, lo usamos.
# Si no, definimos un no-op para que los imports no fallen.
try:
    from empaques.utils_inv import post_ledger as _post_ledger_empaques  # type: ignore
except Exception:
    _post_ledger_empaques = None

def post_ledger(*args, **kwargs):
    """
    Proxy: si existe post_ledger en empaques.utils_inv lo invoca, si no, no-op.
    Mantiene compatibilidad con `from .utils_inv import post_ledger` en arandano.
    """
    if _post_ledger_empaques is not None:
        return _post_ledger_empaques(*args, **kwargs)
    # No-op si aún no está lista la integración
    return None