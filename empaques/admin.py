# empaques/admin.py
from django.contrib import admin
from django import forms

from .models import (
    Shipment, ShipmentItem, Presentation,
    Worker, Field, Crew, CrewMember, AttendanceRecord
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


# -----------------------------
# Personal / Asistencia
# -----------------------------
@admin.register(Worker)
class WorkerAdmin(admin.ModelAdmin):
    list_display = ("full_name", "company", "position", "active", "cycle_started_at")
    list_filter = ("company", "position", "active")
    search_fields = ("full_name",)


@admin.register(Field)
class FieldAdmin(admin.ModelAdmin):
    list_display = ("name",)
    search_fields = ("name",)


class CrewMemberInline(admin.TabularInline):
    model = CrewMember
    extra = 0


@admin.register(Crew)
class CrewAdmin(admin.ModelAdmin):
    list_display = ("name", "field", "leader")
    list_filter = ("field",)
    search_fields = ("name", "leader__full_name")
    inlines = [CrewMemberInline]


@admin.register(AttendanceRecord)
class AttendanceAdmin(admin.ModelAdmin):
    list_display = ("date", "field", "crew", "worker", "status", "observation", "created_by")
    list_filter = ("date", "field", "crew", "status", "worker__company")
    search_fields = ("worker__full_name", "observation")
