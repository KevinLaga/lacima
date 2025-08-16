from django.db import models

class Presentation(models.Model):
    

    # Factor para convertir a “equivalentes de 11 lbs”
    name = models.CharField(max_length=100, unique=True)
    conversion_factor = models.DecimalField(max_digits=32, decimal_places=16)
    price = models.DecimalField(max_digits=6, decimal_places=2)


    def __str__(self):
        # Mostrará algo como “11 lbs – Jumbo”
        return self.name
        


class Shipment(models.Model):
    tracking_number = models.CharField(max_length=30, verbose_name="Número de orden")  
    date = models.DateField(verbose_name="Fecha")  
    carrier = models.CharField(max_length=50, verbose_name="Transportista", blank=True)
    tractor_plates = models.CharField(max_length=20, verbose_name="Placas tractor", blank=True)
    box_plates = models.CharField(max_length=20, verbose_name="Placas caja", blank=True)
    driver = models.CharField(max_length=50, verbose_name="Operador", blank=True)
    departure_time = models.CharField(max_length=20, verbose_name="Horario de salida", blank=True)
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
            ('export_reports', 'Puede exportar reportes'),
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

