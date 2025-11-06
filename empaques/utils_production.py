# empaques/utils_production.py
from datetime import date
from django.db.models import Sum
from django.utils import timezone

# Ajusta esto a tu definición real de temporada
def get_season_bounds(hoy=None):
    # Ejemplo: temporada 2025 inicia el 1 de agosto 2025
    # ⚠️ cámbialo por tu lógica/tabla Season si la tienes.
    hoy = hoy or timezone.localdate()
    season_start = date(hoy.year if hoy.month >= 8 else hoy.year - 1, 8, 1)
    return season_start, hoy

def build_production_context(hasta=None):
    """
    Devuelve un diccionario con:
      - acumulados de temporada (campo y empacado)
      - totales del día (si quieres)
      - filas para tabla/pantalla
    """
    from .models import ProductionDaily  # o tus modelos reales: Harvest, Packing, etc.

    hoy_local = timezone.localdate()
    hasta = hasta or hoy_local
    season_start, _ = get_season_bounds(hoy_local)

    # IMPORTANTE: usa __date si tus campos son DateTime, o directamente el DateField
    qs_temp = ProductionDaily.objects.filter(fecha__range=(season_start, hasta))

    acum_campo = (qs_temp.aggregate(s=Sum('cajas_campo'))['s'] or 0)
    acum_emp   = (qs_temp.aggregate(s=Sum('cajas_empacadas'))['s'] or 0)

    # Datos para la tabla de la vista (si procede)
    # Si quieres también mostrar el día actual
    qs_dia = ProductionDaily.objects.filter(fecha=hasta).order_by('presentacion', 'tamanio')

    return {
        'hasta': hasta,
        'season_start': season_start,
        'acum_cajas_campo': int(acum_campo),
        'acum_cajas_empacadas': int(acum_emp),
        'rows_dia': list(qs_dia),
    }