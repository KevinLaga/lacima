# arandano/models.py
from django.db import models
from django.conf import settings
from django.contrib.contenttypes.fields import GenericForeignKey
from django.contrib.contenttypes.models import ContentType
from decimal import Decimal

DEC_KG = (12, 5)  # 5 decimales
class DestinoSalida(models.TextChoices):
    EMPAQUE = "EMPAQUE", "Empaque"
    OTRO    = "OTRO",    "Otro destino"

class Campo(models.Model):
    nombre = models.CharField(max_length=100, unique=True)
    activo = models.BooleanField(default=True)
    def __str__(self): return self.nombre

class Variedad(models.Model):
    nombre = models.CharField(max_length=100, unique=True)
    activo = models.BooleanField(default=True)
    def __str__(self): return self.nombre

class CampoVariedad(models.Model):
    campo = models.ForeignKey(Campo, on_delete=models.CASCADE)
    variedad = models.ForeignKey(Variedad, on_delete=models.CASCADE)
    activo = models.BooleanField(default=True)
    class Meta:
        unique_together = ("campo", "variedad")

# ===== Producción en KG =====

class ProduccionDia(models.Model):
    fecha = models.DateField()
    campo = models.ForeignKey(Campo, on_delete=models.PROTECT, related_name="producciones")
    rezaga_kg = models.DecimalField(max_digits=10, decimal_places=2, default=0)  # rezaga en KG
    notas = models.TextField(blank=True)
    creado = models.DateTimeField(auto_now_add=True)
    actualizado = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ("fecha", "campo")
        ordering = ["-fecha", "-id"]
        verbose_name = "Producción diaria"
        verbose_name_plural = "Producciones diarias"

    def __str__(self):
        return f"{self.fecha} · {self.campo}"

    @property
    def total_general_kg(self):
        # suma de kg en todos los renglones
        return sum(self.items.values_list("kg", flat=True)) or 0

    @property
    def empacado_kg(self):
        tg = self.total_general_kg or 0
        rz = self.rezaga_kg or 0
        return max(tg - rz, 0)

    @property
    def porcentaje_rezaga(self):
        tg = float(self.total_general_kg or 0)
        if tg <= 0:
            return 0.0
        return float(self.rezaga_kg or 0) * 100.0 / tg

class ProduccionItem(models.Model):
    produccion = models.ForeignKey(ProduccionDia, related_name="items", on_delete=models.CASCADE)
    variedad = models.ForeignKey(Variedad, on_delete=models.PROTECT)
    kg = models.DecimalField(max_digits=10, decimal_places=5, null=True, blank=True)  # solo KG
    # NUEVO: clamshells por variedad
    cs_6oz   = models.PositiveIntegerField(default=0)
    cs_9_8oz = models.PositiveIntegerField(default=0)
    cs_18oz  = models.PositiveIntegerField(default=0)

    class Meta:
        unique_together = ("produccion", "variedad")

    def __str__(self):
        return f"{self.produccion.campo} • {self.variedad}"
    
class SalidaDia(models.Model):
    fecha = models.DateField()
    # antes: related_name="salidas"
    campo = models.ForeignKey(Campo, on_delete=models.PROTECT, related_name="salidas_dia")
    notas = models.TextField(blank=True)
    destino = models.CharField(
        max_length=16,
        choices=DestinoSalida.choices,
        default=DestinoSalida.EMPAQUE,
    )
    destino_detalle = models.CharField(max_length=120, blank=True, default="")
    
    

    creado_en = models.DateTimeField(auto_now_add=True)


    class Meta:
        ordering = ["-fecha", "-id"]
        verbose_name = "Salida (día)"
        verbose_name_plural = "Salidas (días)"

    def __str__(self):
        return f"{self.fecha} · {self.campo}"


class SalidaItem(models.Model):
    salida = models.ForeignKey(SalidaDia, on_delete=models.CASCADE, related_name="items")
    variedad = models.ForeignKey(Variedad, on_delete=models.PROTECT)
    kg = models.DecimalField(max_digits=12, decimal_places=5, default=0)
    # NUEVO: clamshells en salida
    cs_6oz   = models.PositiveIntegerField(default=0)
    cs_9_8oz = models.PositiveIntegerField(default=0)
    cs_18oz  = models.PositiveIntegerField(default=0)

    class Meta:
        verbose_name = "Renglón de salida"
        verbose_name_plural = "Renglones de salida"

    def __str__(self):
        return f"{self.variedad} – {self.kg} kg"
    
# ===== Ledger / Inventario =====

class ScopeInv(models.TextChoices):
    CAMPO   = "CAMPO",   "Campo"
    EMPAQUE = "EMPAQUE", "Empaque"

class ConceptoInv(models.TextChoices):
    ENTRADA_CAMPO                = "ENTRADA_CAMPO", "Entrada por Producción (Campo)"
    SALIDA_CAMPO_OTRO            = "SALIDA_CAMPO_OTRO", "Salida Campo (Otro destino)"
    TRANSFER_CAMPO_A_EMPAQUE     = "TRANSFER_CAMPO_A_EMPAQUE", "Transferencia Campo→Empaque"
    ENTRADA_EMPAQUE_DESDE_CAMPO  = "ENTRADA_EMPAQUE_DESDE_CAMPO", "Entrada Empaque desde Campo"
    SALIDA_EMPAQUE_EMBARQUE      = "SALIDA_EMPAQUE_EMBARQUE", "Salida Empaque por Embarque"
    AJUSTE                       = "AJUSTE", "Ajuste manual"

class InventarioMovimiento(models.Model):
    # Dónde impacta el stock
    scope    = models.CharField(max_length=10, choices=ScopeInv.choices)
    campo    = models.ForeignKey('Campo', null=True, blank=True, on_delete=models.PROTECT,
                                 related_name="movimientos_inv")
    variedad = models.ForeignKey('Variedad', on_delete=models.PROTECT, related_name="movimientos_inv")

    fecha    = models.DateField()
    concepto = models.CharField(max_length=40, choices=ConceptoInv.choices)

    kg       = models.DecimalField(max_digits=DEC_KG[0], decimal_places=DEC_KG[1], default=0)
    cs_6oz   = models.PositiveIntegerField(default=0)
    cs_9_8oz = models.PositiveIntegerField(default=0)
    cs_18oz  = models.PositiveIntegerField(default=0)

    notas    = models.CharField(max_length=200, blank=True)

    # Referencia al documento origen (ProduccionDia / SalidaDia / Embarque / etc.)
    content_type = models.ForeignKey(ContentType, on_delete=models.SET_NULL, null=True, blank=True)
    object_id    = models.PositiveIntegerField(null=True, blank=True)
    ref          = GenericForeignKey('content_type', 'object_id')

    creado_en    = models.DateTimeField(auto_now_add=True)

    class Meta:
        indexes = [
            models.Index(fields=["scope", "fecha"]),
            models.Index(fields=["scope", "variedad", "fecha"]),
            models.Index(fields=["scope", "campo", "variedad", "fecha"]),
        ]
        ordering = ["fecha", "id"]

    def __str__(self):
        cm = f"{self.campo} · " if self.campo else ""
        return f"[{self.scope}] {cm}{self.variedad} · {self.concepto} · {self.kg} kg"

# --- Salida "simple" por movimiento ---
class Salida(models.Model):
    fecha = models.DateField()
    # antes: related_name="salidas"
    campo = models.ForeignKey(Campo, on_delete=models.PROTECT, related_name="salidas_mov")
    # mejor evitar colisión también en Variedad
    variedad = models.ForeignKey(Variedad, on_delete=models.PROTECT, null=True, blank=True, related_name="salidas_mov")
    kg = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    notas = models.CharField(max_length=200, blank=True)

    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-fecha", "-id"]
        verbose_name = "Salida"
        verbose_name_plural = "Salidas"

    def __str__(self):
        v = f" · {self.variedad}" if self.variedad else ""
        return f"{self.fecha} · {self.campo}{v} · {self.kg} kg"
class InventoryLedger(models.Model):
    fecha     = models.DateField()
    scope     = models.CharField(max_length=10, choices=ScopeInv.choices)
    variedad  = models.ForeignKey(Variedad, on_delete=models.PROTECT)
    campo     = models.ForeignKey(Campo, null=True, blank=True, on_delete=models.PROTECT)
    # signo + para entradas, - para salidas
    kg        = models.DecimalField(max_digits=14, decimal_places=5, default=Decimal("0"))
    cs_6oz    = models.IntegerField(default=0)
    cs_9_8oz  = models.IntegerField(default=0)
    cs_18oz   = models.IntegerField(default=0)
    ref_app   = models.CharField(max_length=32, blank=True)  # "PROD", "SALIDA", "MOV"
    ref_id    = models.CharField(max_length=64, blank=True)  # id del objeto origen
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        indexes = [
            models.Index(fields=["scope", "variedad", "fecha"]),
            models.Index(fields=["scope", "variedad", "campo", "fecha"]),
        ]
        ordering = ["fecha", "id"]

    def __str__(self):
        c = f" · {self.campo}" if self.campo_id else ""
        return f"{self.fecha} · {self.scope}{c} · {self.variedad} · {self.kg:+.5f}kg"