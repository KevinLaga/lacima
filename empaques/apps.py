# empaques/apps.py
from django.apps import AppConfig

class EmpaquesConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "empaques"

    def ready(self):
        """
        Asegura que exista el grupo 'capturista' con permisos de captura mínima.
        Se ignoran errores cuando la BD aún no está migrada (primera corrida).
        """
        try:
            from django.contrib.auth.models import Group, Permission
            from django.db.utils import OperationalError, ProgrammingError

            # Crea/obtiene el grupo
            capturista, _ = Group.objects.get_or_create(name="capturista")

            # Permisos mínimos: solo "add" de Shipment y ShipmentItem
            wanted = [
                ("empaques", "shipment", "add_shipment"),
                ("empaques", "shipmentitem", "add_shipmentitem"),
            ]

            for app_label, model, codename in wanted:
                try:
                    perm = Permission.objects.get(
                        content_type__app_label=app_label,
                        content_type__model=model,
                        codename=codename,
                    )
                    capturista.permissions.add(perm)
                except Permission.DoesNotExist:
                    # Si aún no existen (por orden de migraciones), lo ignoramos.
                    pass

        except (OperationalError, ProgrammingError):
            # Migraciones aún no listas; no hacer nada.
            pass
