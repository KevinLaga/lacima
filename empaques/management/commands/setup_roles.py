from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group, Permission
from django.contrib.contenttypes.models import ContentType
from empaques.models import Shipment, ShipmentItem, Presentation

class Command(BaseCommand):
    help = "Crea grupos y asigna permisos (Capturista, Gestor)."

    def handle(self, *args, **options):
        ct_shipment = ContentType.objects.get_for_model(Shipment)
        ct_item     = ContentType.objects.get_for_model(ShipmentItem)
        ct_pres     = ContentType.objects.get_for_model(Presentation)

        add_shipment    = Permission.objects.get(codename='add_shipment', content_type=ct_shipment)
        change_shipment = Permission.objects.get(codename='change_shipment', content_type=ct_shipment)
        # OJO: NO damos view_shipment al Capturista

        add_item    = Permission.objects.get(codename='add_shipmentitem', content_type=ct_item)
        change_item = Permission.objects.get(codename='change_shipmentitem', content_type=ct_item)

        view_pres   = Permission.objects.get(codename='view_presentation', content_type=ct_pres)

        export_reports = Permission.objects.get(codename='export_reports', content_type=ct_shipment)

        # --- Capturista: solo captura (NO ve listas/reportes) ---
        capturista, _ = Group.objects.get_or_create(name='Capturista')
        capturista.permissions.set([
            add_shipment, change_shipment,   # crear/editar durante la captura
            add_item, change_item,           # crear/editar renglones
            view_pres,                       # ver presentaciones en el form
        ])
        capturista.save()

        # --- Gestor: puede ver, exportar y gestionar presentaciones ---
        view_shipment  = Permission.objects.get(codename='view_shipment', content_type=ct_shipment)
        view_item      = Permission.objects.get(codename='view_shipmentitem', content_type=ct_item)
        add_pres       = Permission.objects.get(codename='add_presentation', content_type=ct_pres)
        change_pres    = Permission.objects.get(codename='change_presentation', content_type=ct_pres)
        delete_pres    = Permission.objects.get(codename='delete_presentation', content_type=ct_pres)

        gestor, _ = Group.objects.get_or_create(name='Gestor')
        gestor.permissions.set([
            view_shipment, add_shipment, change_shipment,
            view_item, add_item, change_item,
            view_pres, add_pres, change_pres, delete_pres,
            export_reports,
        ])
        gestor.save()

        self.stdout.write(self.style.SUCCESS("Grupos y permisos configurados."))
