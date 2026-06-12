# empaques/urls_inventory.py
from django.urls import path
from . import views_inventory as v

urlpatterns = [
    # ── Catálogo y stock ──────────────────────────────────────────────
    path("almacen/", v.almacen_list, name="almacen_list"),
    path("almacen/articulo/nuevo/", v.almacen_item_new, name="almacen_item_new"),
    path("almacen/historial/<int:item_id>/", v.almacen_kardex, name="almacen_kardex"),

    # ── Pedimentos ───────────────────────────────────────────────────
    path("almacen/pedimentos/", v.pedimento_list, name="pedimento_list"),
    path("almacen/pedimentos/nuevo/", v.pedimento_new, name="pedimento_new"),
    path("almacen/pedimentos/<int:pk>/", v.pedimento_detail, name="pedimento_detail"),

    # ── Remisiones ───────────────────────────────────────────────────
    path("almacen/remisiones/", v.remision_list, name="remision_list"),
    path("almacen/remisiones/nueva/", v.remision_new, name="remision_new"),
    path("almacen/remisiones/<int:pk>/", v.remision_detail, name="remision_detail"),

    # ── Inventario inicial ───────────────────────────────────────────
    path("almacen/inventario-inicial/", v.inventario_inicial, name="inventario_inicial"),

    # ── Legacy: movimiento individual ────────────────────────────────
    path("almacen/movimiento/nuevo/", v.almacen_movement_new, name="almacen_movement_new"),
    path("almacen/movimiento/nuevo/<str:tipo>/", v.almacen_movement_new, name="almacen_movement_new_tipo"),
]
