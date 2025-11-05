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
DATA_DIR = Path(os.environ.get("DATA_DIR", str(BASE_DIR / "data")))  # <-- usa /srv/lacima/data en prod
DATA_DIR.mkdir(parents=True, exist_ok=True)
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": str(DATA_DIR / "db.sqlite3"),
    }
}

# -----------------------------
# Utilidades
# -----------------------------
def env_bool(name: str, default: bool = False) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return val.lower() in ("1", "true", "yes", "on")

def env_list(name: str, default: str = "") -> list[str]:
    raw = os.getenv(name, default)
    return [x.strip() for x in raw.split(",") if x.strip()]

# -----------------------------
# Debug / Secret / Hosts
# -----------------------------
DEBUG = env_bool("DEBUG", True)  # True en local por defecto
SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-CHANGE-ME")

DEFAULT_HOSTS = "lacima.onrender.com,localhost,127.0.0.1"
ALLOWED_HOSTS = env_list("ALLOWED_HOSTS", DEFAULT_HOSTS)

# CSRF: confía en los hosts (con https) excepto localhost
CSRF_TRUSTED_ORIGINS = [
    f"https://{h.lstrip('.')}" for h in ALLOWED_HOSTS
    if h not in ("localhost", "127.0.0.1")
]

# Render/Proxies: respeta X-Forwarded-Proto
SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")

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
    "empaques.apps.EmpaquesConfig",
]

# -----------------------------
# Middleware (WhiteNoise)
# -----------------------------
MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "whitenoise.middleware.WhiteNoiseMiddleware",
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
        "DIRS": [],  # usa templates de las apps
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
# Base de datos
# - En Render: DATABASE_URL (Postgres)
# - En local: SQLite por defecto (sin SSL)
# -----------------------------
DATABASES = {
    "default": dj_database_url.config(
        env="DATABASE_URL",
        default=f"sqlite:///{BASE_DIR / 'db.sqlite3'}",
        conn_max_age=600,
        ssl_require=not DEBUG,  # solo exige SSL cuando no estás en DEBUG
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
LANGUAGE_CODE = "es-mx"
TIME_ZONE = "America/Mazatlan"
USE_I18N = True
USE_TZ = True

# -----------------------------
# Archivos estáticos (WhiteNoise)
# -----------------------------
STATIC_URL = "/static/"
STATICFILES_DIRS = [BASE_DIR / "static"]  # para tus logos, CSS, etc.
STATIC_ROOT = BASE_DIR / "staticfiles"    # donde collectstatic deja todo
STATICFILES_STORAGE = "whitenoise.storage.CompressedManifestStaticFilesStorage"

# -----------------------------
# Seguridad extra en producción
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
LOGIN_REDIRECT_URL = "/empaques/nuevo/"
LOGOUT_REDIRECT_URL = "/accounts/login/"

# -----------------------------
# Default PK
# -----------------------------
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# -----------------------------
# Carpeta de reportes (si la usas)
# -----------------------------
REPORTS_ROOT = BASE_DIR / "reports"
