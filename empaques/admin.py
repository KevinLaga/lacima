from django.contrib import admin
from django import forms
from .models import Shipment, ShipmentItem, Presentation

class ShipmentAdminForm(forms.ModelForm):
    class Meta:
        model = Shipment
        fields = "__all__"

    def clean_tarimas_peco(self):
        v = self.cleaned_data.get("tarimas_peco")
        if v is not None and v < 0:
            raise forms.ValidationError("Tarimas PECO no puede ser negativo.")
        return v

class ShipmentItemInline(admin.TabularInline):
    model = ShipmentItem
    extra = 0
    fields = ("cliente", "presentation", "size", "quantity", "tarima", "temperatura")
    autocomplete_fields = ("presentation",)

@admin.register(Shipment)
class ShipmentAdmin(admin.ModelAdmin):
    form = ShipmentAdminForm

    list_display = (
        "date", "tracking_number", "carrier",
        "invoice_number", "tarimas_peco",
    )
    list_filter = ("date", "carrier",)
    search_fields = (
        "tracking_number", "invoice_number", "carrier",
        "driver", "tractor_plates", "box_plates",
    )
    date_hierarchy = "date"
    inlines = [ShipmentItemInline]

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
            "fields": ("tarimas_peco",)   # ← aquí aparece y se edita en admin
        }),
    )

@admin.register(Presentation)
class PresentationAdmin(admin.ModelAdmin):
    list_display = ("name", "conversion_factor", "price")
    search_fields = ("name",)

# Si no tenías ShipmentItem registrado aparte:
@admin.register(ShipmentItem)
class ShipmentItemAdmin(admin.ModelAdmin):
    list_display = ("shipment", "cliente", "presentation", "size", "quantity", "tarima", "temperatura")
    search_fields = ("cliente", "shipment__tracking_number")
    autocomplete_fields = ("presentation", "shipment")