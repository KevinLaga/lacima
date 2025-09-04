from django.urls import path, include
from . import views

urlpatterns = [
    path('nuevo/', views.shipment_create, name='shipment_create'),
    path('lista/', views.shipment_list, name='shipment_list'),   
    path('reporte/', views.daily_report, name='daily_report'),
    path('accounts/', include('django.contrib.auth.urls')),
    path('reporte/shipment/<int:shipment_id>/', views.daily_report, name='daily_report_by_id'),
    path('reporte/tracking/<str:tracking>/', views.daily_report, name='daily_report_by_tracking'),

    # Producción (las que sí tenemos implementadas)
    path("produccion/", views.production_today, name="production_today"),
    path("produccion/dias/", views.production_days, name="production_days"),
    path("produccion/<slug:prod_date>/xlsx/", views.production_xlsx, name="production_xlsx"),
    path("produccion/ayer/", views.production_yesterday, name="production_yesterday"),
    
]
