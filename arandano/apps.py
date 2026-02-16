from django.apps import AppConfig


class ArandanoConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'arandano'
    verbose_name = "Arándano"

    
    def ready(self):
        from . import signals  # noqa
