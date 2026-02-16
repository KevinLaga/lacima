from django.urls import path
from . import views

app_name = "arandano"

urlpatterns = [
    path("", views.produccion_list, name="arandano_produccion_list"),
    path("nuevo/", views.produccion_create, name="arandano_produccion_create"),
    path("excel/<int:pk>/", views.produccion_excel_dia, name="arandano_produccion_excel_dia"),
    path("excel/rango/", views.arandano_produccion_excel_rango, name="arandano_produccion_excel_rango"),
    path("salidas/", views.salidas_list, name="salidas_list"),
    path("salidas/nuevo/", views.salidas_create, name="salidas_create"),
    path("salidas/excel/", views.salidas_excel_rango, name="salidas_excel_rango"),
    path("salidas/excel/inventario/", views.salidas_excel_inventario, name="salidas_excel_inventario"),
    
]