from django.urls import path, include
from . import views
from empaques.views import post_login_redirect  # <-- importa la vista

urlpatterns = [
path('nuevo/', views.shipment_create, name='shipment_create'),
    path('lista/', views.shipment_list, name='shipment_list'),   
    path('reporte/', views.daily_report, name='daily_report'),
    path('accounts/', include('django.contrib.auth.urls')),
    path('reporte/shipment/<int:shipment_id>/', views.daily_report, name='daily_report_by_id'),
    path('reporte/tracking/<str:tracking>/', views.daily_report, name='daily_report_by_tracking'),
]
