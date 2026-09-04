from django.urls import path

from . import views

app_name = "creditos"

urlpatterns = [
    path("",                  views.credito_list,   name="credito_list"),
    path("nuevo/",            views.credito_create, name="credito_create"),
    path("<int:pk>/",         views.credito_detail, name="credito_detail"),
    path("<int:pk>/editar/",  views.credito_edit,   name="credito_edit"),
    path("excel/",            views.credito_export_xlsx, name="credito_export_xlsx"),
    path("<int:pk>/plan-excel/", views.credito_plan_xlsx, name="credito_plan_xlsx"),
    path("abono/<int:pk>/eliminar/", views.abono_delete, name="abono_delete"),
    path("garantias/",        views.garantia_list,  name="garantia_list"),
    path("garantias/<int:pk>/eliminar/", views.garantia_delete, name="garantia_delete"),
]
