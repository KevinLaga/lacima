"""
Django settings for lacima project
"""

from pathlib import Path
import os
import dj_database_url

# -----------------------------
# Rutas base
# -----------------------------
BASE_DIR = Path(__file__).resolve().parent.parent

DEBUG = False  # en producción

# PON AQUÍ TU HOST REAL DE RENDER
ALLOWED_HOSTS = ["lacima.onrender.com"]
CSRF_TRUSTED_ORIGINS = ["https://lacima.onrender.com"]

# -----------------------------
# Claves y modo debug
# -----------------------------
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-CHANGE-ME")
DEBUG = os.environ.get("DEBUG", "True") == "True"

# Dominios permitidos
if DEBUG:
    ALLOWED_HOSTS = []
    CSRF_TRUSTED_ORIGINS = []
else:
    # Ajusta estos valores para producción
    ALLOWED_HOSTS = os.environ.get("ALLOWED_HOSTS", "").split(",") if os.environ.get("ALLOWED_HOSTS") else []
    CSRF_TRUSTED_ORIGINS = os.environ.get("CSRF_TRUSTED_ORIGINS", "").split(",") if os.environ.get("CSRF_TRUSTED_ORIGINS") else []

# -----------------------------
# Apps instaladas
# -----------------------------
INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "empaques",
]

# -----------------------------
# Middleware (Whitenoise ya incluido)
# -----------------------------
MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "whitenoise.middleware.WhiteNoiseMiddleware",  # <- importante para estáticos en prod
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

# -----------------------------
# URLs / WSGI
# -----------------------------
ROOT_URLCONF = "lacima.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],  # usa templates de apps (APP_DIRS=True)
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "lacima.wsgi.application"

# -----------------------------
# Base de datos (usa DATABASE_URL si existe; si no, SQLite)
# -----------------------------
DATABASES = {
    "default": dj_database_url.config(
        env="DATABASE_URL",
        default=f"sqlite:///{BASE_DIR / 'db.sqlite3'}",
        conn_max_age=600,
        ssl_require=False if DEBUG else True,
    )
}

# -----------------------------
# Validadores de password
# -----------------------------
AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

# -----------------------------
# Internacionalización / Zona horaria
# -----------------------------
LANGUAGE_CODE = "es"  # interfaz admin en español
TIME_ZONE = "America/Mazatlan"
USE_I18N = True
USE_TZ = True

# -----------------------------
# Archivos estáticos
# -----------------------------
STATIC_URL = "/static/"

# Carpeta donde pones tus archivos estáticos propios (logos, css, etc.)
# Crea esta carpeta si no existe: BASE_DIR / "static"
STATICFILES_DIRS = [
    BASE_DIR / "static",
]

# Carpeta a la que collectstatic copiará todo (no la crees tú)
STATIC_ROOT = BASE_DIR / "staticfiles"

# Whitenoise: servir estáticos en prod
STATICFILES_STORAGE = "whitenoise.storage.CompressedManifestStaticFilesStorage"

# -----------------------------
# Seguridad (solo en producción)
# -----------------------------
if not DEBUG:
    SECURE_SSL_REDIRECT = True
    SESSION_COOKIE_SECURE = True
    CSRF_COOKIE_SECURE = True
    SECURE_HSTS_SECONDS = 60 * 60 * 24 * 30  # 30 días
    SECURE_HSTS_INCLUDE_SUBDOMAINS = True
    SECURE_HSTS_PRELOAD = True

# -----------------------------
# Login/Logout
# -----------------------------
LOGIN_URL = "/accounts/login/"
LOGIN_REDIRECT_URL = "/empaques/lista/"
LOGOUT_REDIRECT_URL = "/accounts/login/"

# -----------------------------
# Default PK
# -----------------------------
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# -----------------------------
# Carpeta de reportes (si la usas)
# -----------------------------
REPORTS_ROOT = BASE_DIR / "reports"