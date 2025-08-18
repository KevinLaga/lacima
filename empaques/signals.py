import os
from django.contrib.auth import get_user_model
from django.db.models.signals import post_migrate
from django.dispatch import receiver

@receiver(post_migrate)
def create_initial_superuser(sender, **kwargs):
    """
    Crea un superusuario si no existe y hay variables de entorno definidas.
    Se ejecuta después de 'migrate'. Idempotente por username.
    """
    username = os.getenv("DJANGO_SUPERUSER_USERNAME")
    email    = os.getenv("DJANGO_SUPERUSER_EMAIL", "")
    password = os.getenv("DJANGO_SUPERUSER_PASSWORD")

    if not username or not password:
        return  # no hay datos -> no hace nada

    User = get_user_model()
    if not User.objects.filter(username=username).exists():
        User.objects.create_superuser(username=username, email=email, password=password)
