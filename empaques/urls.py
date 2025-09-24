from django.urls import path, include
from . import views
from django.http import HttpResponse
from django.db import connection
from django.conf import settings
def _dbg_db(request):
    with connection.cursor() as c:
        # esto devuelve la ruta REAL del archivo que abrió sqlite
        c.execute("PRAGMA database_list;")
        dblist = c.fetchall()
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='empaques_productiondisplay';")
        has = bool(c.fetchone())
    return HttpResponse(
        f"settings.DB_NAME={settings.DATABASES['default']['NAME']}<br>"
        f"PRAGMA database_list={dblist}<br>"
        f"has_empaques_productiondisplay={has}"
    )


urlpatterns = [
    path('nuevo/', views.shipment_create, name='shipment_create'),
    path('lista/', views.shipment_list, name='shipment_list'),   
    path('reporte/', views.daily_report, name='daily_report'),
    path('accounts/', include('django.contrib.auth.urls')),
    path('reporte/shipment/<int:shipment_id>/', views.daily_report, name='daily_report_by_id'),
    path('reporte/tracking/<str:tracking>/', views.daily_report, name='daily_report_by_tracking'),
    path("__debug_db__", _dbg_db),

    # Producción (las que sí tenemos implementadas)
    path("produccion/", views.production_today, name="production_today"),
    path("produccion/dias/", views.production_days, name="production_days"),
    path("produccion/<slug:prod_date>/xlsx/", views.production_xlsx, name="production_xlsx"),
    path("", include("empaques.urls_inventory")),
    path("produccion/ayer/", views.production_yesterday, name="production_yesterday"),
    
]
