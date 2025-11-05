from django.db import models
from django.contrib.auth import get_user_model
from django.db.models import Sum, F, Case, When, Value, DecimalField
from django.db.models.functions import Coalesce
from decimal import Decimal
from django.db.models import Sum, Q, DecimalField, Value as V
from django.contrib.auth.models import User
from django.conf import settings






class Presentation(models.Model):
    

    # Factor para convertir a “equivalentes de 11 lbs”
    name = models.CharField(max_length=100, unique=True)
    conversion_factor = models.DecimalField(max_digits=32, decimal_places=16)
    price = models.DecimalField(max_digits=32, decimal_places=16)


    def __str__(self):
        # Mostrará algo como “11 lbs – Jumbo”
        return self.name
        


class Shipment(models.Model):
    tarimas_peco = models.PositiveIntegerField(
        "Tarimas PECO",
        blank=True,
        null=True,
        help_text="Número de tarimas tipo PECO en el embarque"
    )
    tracking_number = models.CharField(max_length=30, verbose_name="Número de orden")  
    order_lacima    = models.CharField("Núm. orden CIMA",        max_length=50, blank=True, null=True)
    order_rc        = models.CharField("Núm. orden RC Organics",  max_length=50, blank=True, null=True)
    order_gourmet   = models.CharField("Núm. orden Gourmet Baja", max_length=50, blank=True, null=True)
    order_gbf       = models.CharField("Núm. orden GBF Farms",    max_length=50, blank=True, null=True)
    order_gh       = models.CharField("Núm. orden GH Farms", max_length=50, blank=True, null=True)
    order_dhg       = models.CharField("Núm. orden AGRICOLA DH & G", max_length=50, blank=True, null=True)
    date = models.DateField(verbose_name="Fecha")  
    carrier = models.CharField(max_length=50, verbose_name="Transportista", blank=True)
    tractor_plates = models.CharField(max_length=20, verbose_name="Placas tractor", blank=True)
    box_plates = models.CharField(max_length=20, verbose_name="Placas caja", blank=True)
    driver = models.CharField(max_length=50, verbose_name="Operador", blank=True)
    departure_time = models.CharField(max_length=20, verbose_name="Horario de salida", blank=True, null=True)
    box = models.CharField(max_length=30, verbose_name="Caja", blank=True)
    box_conditions = models.CharField(max_length=100, verbose_name="Condiciones de la caja", blank=True)
    box_free_of_odors = models.CharField(max_length=50, verbose_name="Caja libre de olores", blank=True)
    ryan = models.CharField(max_length=50, verbose_name="Ryan", blank=True)
    seal_1 = models.CharField(max_length=30, verbose_name="Sello 1", blank=True)
    seal_2 = models.CharField(max_length=30, verbose_name="Sello 2", blank=True)
    seal_3 = models.CharField(max_length=30, verbose_name="Sello 3", blank=True)
    seal_4 = models.CharField(max_length=30, verbose_name="Sello 4", blank=True)
    chismografo = models.CharField(max_length=50, verbose_name="Chismógrafo", blank=True)
    delivery_signature = models.CharField(max_length=50, verbose_name="Firma de entrega", blank=True)
    driver_signature = models.CharField(max_length=50, verbose_name="Firma de operador", blank=True)
    invoice_number = models.CharField(max_length=30, verbose_name="Número de factura", blank=True)  
    class Meta:
        permissions = [
            ("can_download_reports", "Puede descargar/exportar reportes"),

        ]


    def __str__(self):
        return f"Embarque {self.tracking_number} – {self.date}"

    @property
    def total_boxes(self):
        # Suma de todas las cajas enviadas en este embarque
        return sum(item.quantity for item in self.items.all())

    @property
    def total_equivalent_11lbs(self):
        # Suma de quantity × conversion_factor de cada item
        return sum(item.quantity * item.presentation.conversion_factor
                   for item in self.items.all())

    @property
    def total_amount(self):
        # Suma de quantity × price de cada item
        return sum(item.quantity * item.presentation.price
                   for item in self.items.all())


SIZE_CHOICES = [
    ('Jumbo',    'Jumbo'),
    ('XLarge',   'X-Large'),
    ('Large',    'Large'),
    ('Standard', 'Standard'),
    ('Small',    'Small'),
    ('Tips',     'Tips'),
]
CLIENTE_CHOICES = [
    ('La Cima Produce', 'La Cima Produce'),
    ('RC Organics', 'RC Organics'),
    ('GH Farms', 'GH Farms'),
    ('Gourmet Baja Farms', 'Gourmet Baja Farms'),
    ('GBF Farms', 'GBF Farms'),
    ('AGRICOLA DH & G', 'AGRICOLA DH & G'),

]
class ShipmentItem(models.Model):
    shipment = models.ForeignKey(Shipment, related_name='items', on_delete=models.CASCADE)
    tarima = models.PositiveIntegerField("Tarima", default=1)
    presentation = models.ForeignKey(Presentation, on_delete=models.PROTECT)
    size = models.CharField(max_length=20, choices=SIZE_CHOICES)
    quantity = models.PositiveIntegerField()
    cliente = models.CharField(
        max_length=32,
        choices=CLIENTE_CHOICES,
        blank=True,    # <---- Permite dejar vacío en formularios
        null=True      # <---- Permite que sea NULL en la base de datos
    )
    temperatura = models.DecimalField(
        max_digits=5,
        decimal_places=1,
        blank=True,
        null=True,
        help_text="Temperatura registrada (puede quedar vacío)"
    )



    
    presentation = models.ForeignKey(Presentation, on_delete=models.PROTECT)

    # —> Esto es el campo SIZE en ShipmentItem:
    SIZE_CHOICES = [
        ('Jumbo',    'Jumbo'),
        ('XLarge',   'X-Large'),
        ('Large',    'Large'),
        ('Standard', 'Standard'),
        ('Small',    'Small'),
        ('Tips',     'Tips'),
    ]
    size = models.CharField(
        max_length=20,
        choices=SIZE_CHOICES,
        default='Standard',
        help_text="Tamaño de esta línea del embarque"
    )

    quantity = models.PositiveIntegerField(
        help_text="Número de cajas enviadas de esta presentación"
    )

# ===== ALMACÉN (top-level, fuera de Shipment/ShipmentItem) =====
from django.core.validators import MinValueValidator
class InventoryItem(models.Model):
    sku = models.CharField("Id", max_length=32, unique=True, blank=True)  # ← etiqueta "Id"
    name = models.CharField(max_length=200)
    location = models.CharField(max_length=120, blank=True)
    unit = models.CharField(max_length=20, default="pz")
    min_stock = models.DecimalField(
        max_digits=12,
        decimal_places=2,
        default=Decimal("0"),      # siempre arranca en 0
        validators=[MinValueValidator(0)],
        blank=True,                 # no obligatorio en formularios
    )
    def save(self, *args, **kwargs):
        # fuerza 0 si viene vacío/negativo
        if self.min_stock is None or self.min_stock < 0:
            self.min_stock = Decimal("0")
        super().save(*args, **kwargs)

    created_at = models.DateTimeField(auto_now_add=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="inventory_items"
    )

    def __str__(self):
        return f"[{self.sku or 'SIN-ID'}] {self.name}"

    def save(self, *args, **kwargs):
        creating = self.pk is None
        super().save(*args, **kwargs)
        if creating and not self.sku:
            self.sku = f"I{self.pk:05d}"  # p.ej. I00001, I00002...
            super().save(update_fields=["sku"])

    @property
    def stock(self):
        """
        Stock actual = Entradas + Ajustes - Salidas
        (cálculo puntual; para listar muchos, usa annotate en la vista)
        """
        agg = self.movements.aggregate(
            ent=Coalesce(
                Sum('quantity', filter=Q(type='IN'), output_field=DecimalField(max_digits=12, decimal_places=2)),
                0, output_field=DecimalField(max_digits=12, decimal_places=2)
            ),
            sal=Coalesce(
                Sum('quantity', filter=Q(type='OUT'), output_field=DecimalField(max_digits=12, decimal_places=2)),
                0, output_field=DecimalField(max_digits=12, decimal_places=2)
            ),
            adj=Coalesce(
                Sum('quantity', filter=Q(type='ADJ'), output_field=DecimalField(max_digits=12, decimal_places=2)),
                0, output_field=DecimalField(max_digits=12, decimal_places=2)
            ),
        )
        return (agg['ent'] or 0) - (agg['sal'] or 0) + (agg['adj'] or 0)


class InventoryMovement(models.Model):
    TYPE_CHOICES = (
        ('IN', 'Entrada'),
        ('OUT', 'Salida'),
        ('ADJ', 'Ajuste'),
    )
    item = models.ForeignKey(InventoryItem, related_name="movements", on_delete=models.PROTECT)
    date = models.DateField()
    type = models.CharField(max_length=3, choices=TYPE_CHOICES)
    quantity = models.DecimalField(max_digits=12, decimal_places=2)
    reference = models.CharField(max_length=120, blank=True)
    notes = models.TextField(blank=True)
    created_by = models.ForeignKey(User, on_delete=models.PROTECT, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ("date", "id")

    def __str__(self):
        return f"{self.get_type_display()} {self.quantity} {self.item.unit} → {self.item}"

    def clean(self):
        from django.core.exceptions import ValidationError
        if self.quantity is None or self.quantity <= 0:
            raise ValidationError("La cantidad debe ser positiva.")
        if self.type == "OUT":
            current = self.item.stock
            if current - self.quantity < 0:
                raise ValidationError(f"Stock insuficiente. Disponible: {current} {self.item.unit}.")
            

# empaques/models.py
from django.db import models

# al inicio del archivo ya debes tener: from django.db import models
# y Presentation definido

class ProductionDisplay(models.Model):
    """
    Catálogo editable desde Admin: define qué (presentación + tamaño)
    aparecen en Producción diaria y en qué orden.
    """
    presentation = models.ForeignKey(
        Presentation,
        on_delete=models.CASCADE,
        related_name="production_displays",
    )
    size = models.CharField(
        max_length=50,
        help_text="Tamaño EXACTO como se captura en los ítems (p. ej. 'Jumbo')."
    )
    order = models.PositiveIntegerField(default=0)
    is_active = models.BooleanField(default=True)

    class Meta:
        unique_together = ('presentation', 'size')
        ordering = ('order', 'presentation__name', 'size')
        verbose_name = "Presentación en Producción"
        verbose_name_plural = "Presentaciones en Producción"

    def __str__(self):
        return f"{self.presentation.name} — {self.size}"
