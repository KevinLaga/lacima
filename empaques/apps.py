from django.apps import AppConfig

class EmpaquesConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "empaques"

    def ready(self):
        # Importa las señales para que se conecten
        from . import signals  # noqa: F401
