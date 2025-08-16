from django.contrib import admin
from .models import Presentation, Shipment, ShipmentItem

class ShipmentItemInline(admin.TabularInline):
    model = ShipmentItem
    extra = 1

class ShipmentAdmin(admin.ModelAdmin):
    list_display = ('tracking_number', 'invoice_number', 'date')
    inlines = [ShipmentItemInline]

@admin.register(Presentation)
class PresentationAdmin(admin.ModelAdmin):
    list_display = ('name', 'conversion_factor', 'price')

admin.site.register(Shipment, ShipmentAdmin)
