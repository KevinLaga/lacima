# empaques/admin.py
from django.contrib import admin
from django import forms
from django.contrib.admin.sites import NotRegistered

from .models import (
    Shipment, ShipmentItem, Presentation,
    InventoryItem, InventoryMovement,
)

# -----------------------------
# Forms
# -----------------------------
class ShipmentAdminForm(forms.ModelForm):
    class Meta:
        model = Shipment
        fields = "__all__"

    def clean_tarimas_peco(self):
        v = self.cleaned_data.get("tarimas_peco")
        if v is not None and v < 0:
            raise forms.ValidationError("Tarimas PECO no puede ser negativo.")
        return v


# -----------------------------
# Inlines
# -----------------------------
class ShipmentItemInline(admin.TabularInline):
    model = ShipmentItem
    extra = 0
    fields = ("cliente", "presentation", "size", "quantity", "tarima", "temperatura")
    autocomplete_fields = ("presentation",)


# -----------------------------
# Shipment
# -----------------------------
@admin.register(Shipment)
class ShipmentAdmin(admin.ModelAdmin):
    form = ShipmentAdminForm
    date_hierarchy = "date"
    list_per_page = 50

    # Muestra y permite editar EN LISTA los números por cliente
    list_display = (
        "id", "date", "tracking_number", "invoice_number",
        "order_lacima", "order_rc", "order_gh", "order_gourmet", "order_gbf",
        "carrier", "tarimas_peco",
    )
    list_display_links = ("id", "tracking_number")  # estos NO pueden ir en list_editable
    list_editable = ("order_lacima", "order_rc", "order_gh", "order_gourmet", "order_gbf")

    list_filter = ("date", "carrier")
    search_fields = (
        "tracking_number", "invoice_number",
        "order_lacima", "order_rc", "order_gh", "order_gourmet", "order_gbf",
        "driver", "carrier", "tractor_plates", "box_plates",
    )

    fieldsets = (
        ("Datos básicos", {
            "fields": ("date", "tracking_number", "invoice_number")
        }),
        ("Transporte", {
            "fields": ("carrier", "driver", "tractor_plates", "box_plates", "box", "departure_time")
        }),
        ("Condiciones", {
            "fields": (
                "box_conditions", "box_free_of_odors", "ryan", "chismografo",
                "seal_1", "seal_2", "seal_3", "seal_4",
            )
        }),
        ("Firmas", {
            "fields": ("delivery_signature", "driver_signature")
        }),
        ("Extras", {
            "fields": ("tarimas_peco",)
        }),
        ("Números de orden por cliente (solo para mostrar en los Excel por cliente)", {
            "fields": ("order_lacima", "order_rc", "order_gh", "order_gourmet", "order_gbf"),
        }),
    )

    inlines = [ShipmentItemInline]


# -----------------------------
# Presentation
# -----------------------------
@admin.register(Presentation)
class PresentationAdmin(admin.ModelAdmin):
    list_display = ("name", "conversion_factor", "price")
    search_fields = ("name",)
    ordering = ("name",)


# -----------------------------
# ShipmentItem (admin directo, además del inline)
# -----------------------------
@admin.register(ShipmentItem)
class ShipmentItemAdmin(admin.ModelAdmin):
    list_display = ("shipment", "cliente", "presentation", "size", "quantity", "tarima", "temperatura")
    search_fields = ("cliente", "shipment__tracking_number", "presentation__name")
    list_filter = ("cliente",)
    autocomplete_fields = ("presentation", "shipment")

from .models import InventoryItem, InventoryMovement

# ---------------------------
#  Inventario (Almacén)
# ---------------------------

# Si ya estaba registrado en otro lado, lo desregistramos para
# volver a registrarlo con esta configuración.
try:
    admin.site.unregister(InventoryItem)
except NotRegistered:
    pass

@admin.register(InventoryItem)
class InventoryItemAdmin(admin.ModelAdmin):
    list_display = ("sku", "name", "location", "unit", "min_stock", "stock_admin", "created_at")
    search_fields = ("sku", "name", "location")
    list_filter = ("unit",)
    readonly_fields = ("created_at",)

    def stock_admin(self, obj):
        # Usa la propiedad .stock del modelo (o cámbialo si prefieres un annotate)
        return obj.stock
    stock_admin.short_description = "Stock"

try:
    admin.site.unregister(InventoryMovement)
except NotRegistered:
    pass

@admin.register(InventoryMovement)
class InventoryMovementAdmin(admin.ModelAdmin):
    list_display = ("date", "item", "type", "quantity", "reference", "created_by", "created_at")
    list_filter  = ("type", "date")
    search_fields = ("item__sku", "item__name", "reference", "notes")
    autocomplete_fields = ("item",)
    readonly_fields = ("created_at",)

from django.contrib import admin
from .models import ProductionDisplay






from django import forms


from .models import ProductionDisplay, ShipmentItem

def _distinct_sizes():
    # Toma los tamaños existentes en items, sin nulos/blank, únicos y ordenados
    return list(
        ShipmentItem.objects
        .exclude(size__isnull=True)
        .exclude(size__exact="")
        .values_list("size", flat=True)
        .distinct()
        .order_by("size")
    )

class ProductionDisplayAdminForm(forms.ModelForm):
    # Sobrescribimos el field para volverlo un ChoiceField
    size = forms.ChoiceField(label="Tamaño", choices=(), required=True)

    class Meta:
        model = ProductionDisplay
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        sizes = _distinct_sizes()

        # Si estás editando y el valor actual no está en la lista (por ejemplo, legacy),
        # lo añadimos al principio para no bloquear la edición.
        current = (self.instance.size or "").strip()
        if current and current not in sizes:
            sizes = [current] + sizes

        self.fields["size"].choices = [(s, s) for s in sizes]

    def clean_size(self):
        # Normaliza espacios
        return (self.cleaned_data.get("size") or "").strip()
    
@admin.register(ProductionDisplay)
class ProductionDisplayAdmin(admin.ModelAdmin):
    form = ProductionDisplayAdminForm
    list_display  = ("presentation", "size", "order", "is_active")
    list_filter   = ("is_active", "presentation")
    search_fields = ("presentation__name", "size")
    ordering      = ("order", "presentation__name", "size")
    autocomplete_fields = ("presentation",)





