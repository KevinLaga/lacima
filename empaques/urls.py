from django.urls import path, include
from . import views

urlpatterns = [
path('nuevo/', views.shipment_create, name='shipment_create'),
    path('lista/', views.shipment_list, name='shipment_list'),   
    path('reporte/', views.daily_report, name='daily_report'),
    path('accounts/', include('django.contrib.auth.urls')),
]
