from django.contrib import admin

from .models import Abono, Credito, Garantia


class AbonoInline(admin.TabularInline):
    model = Abono
    extra = 1
    fields = ("fecha", "monto", "referencia", "nota")


@admin.register(Garantia)
class GarantiaAdmin(admin.ModelAdmin):
    list_display  = ("nombre", "descripcion", "precio", "activo")
    list_editable = ("precio", "activo")
    list_filter   = ("activo",)
    search_fields = ("nombre", "descripcion")


@admin.register(Credito)
class CreditoAdmin(admin.ModelAdmin):
    list_display = (
        "id", "empresa", "banco", "tipo_credito_label", "garantia",
        "moneda", "monto", "col_abonado", "col_saldo",
        "fecha_disposicion", "fecha_vencimiento", "col_estado",
    )
    list_filter   = ("empresa", "banco", "tipo_credito", "moneda", "fecha_vencimiento")
    search_fields = ("referencia", "notas", "tipo_otro", "garantia__nombre")
    date_hierarchy = "fecha_vencimiento"
    inlines = [AbonoInline]

    fieldsets = (
        ("Empresa y banco", {
            "fields": ("empresa", "banco", "referencia"),
        }),
        ("Condiciones del crédito", {
            "fields": ("tipo_credito", "tipo_otro", "garantia",
                       "moneda", "monto", "tasa", "plazo_meses"),
        }),
        ("Fechas", {
            "fields": ("fecha_disposicion", "fecha_vencimiento"),
        }),
        ("Otros", {
            "fields": ("notas",),
        }),
    )

    @admin.display(description="Tipo")
    def tipo_credito_label(self, obj):
        return obj.tipo_credito_label

    @admin.display(description="Abonado")
    def col_abonado(self, obj):
        return obj.abonado_fmt

    @admin.display(description="Saldo")
    def col_saldo(self, obj):
        return obj.saldo_fmt

    @admin.display(description="Estado")
    def col_estado(self, obj):
        return obj.estado_label


@admin.register(Abono)
class AbonoAdmin(admin.ModelAdmin):
    list_display  = ("id", "credito", "fecha", "monto", "referencia")
    list_filter   = ("fecha", "credito__empresa", "credito__banco")
    search_fields = ("referencia", "nota")
    date_hierarchy = "fecha"
