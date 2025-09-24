# empaques/urls_inventory.py
from django.urls import path
from . import views_inventory as v

urlpatterns = [
    path("almacen/", v.almacen_list, name="almacen_list"),
    path("almacen/nuevo/", v.almacen_item_new, name="almacen_item_new"),
    path("almacen/movimiento/nuevo/", v.almacen_movement_new, name="almacen_movement_new"),
    path("almacen/movimiento/nuevo/<str:tipo>/", v.almacen_movement_new, name="almacen_movement_new_tipo"),
    path("almacen/historial/<int:item_id>/", v.almacen_kardex, name="almacen_kardex"),  # “Historial de movimientos”
]