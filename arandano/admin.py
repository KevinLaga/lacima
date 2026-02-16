from django.contrib import admin
from .models import Campo, Variedad, CampoVariedad, ProduccionDia, ProduccionItem, SalidaDia, SalidaItem, Salida

@admin.register(Campo)
class CampoAdmin(admin.ModelAdmin):
    list_display  = ("nombre", "activo")
    list_filter   = ("activo",)
    search_fields = ("nombre",)
    ordering      = ("nombre",)

@admin.register(Variedad)
class VariedadAdmin(admin.ModelAdmin):
    list_display  = ("nombre", "activo")
    list_filter   = ("activo",)
    search_fields = ("nombre",)
    ordering      = ("nombre",)

@admin.register(CampoVariedad)
class CampoVariedadAdmin(admin.ModelAdmin):
    list_display  = ("campo", "variedad", "activo")
    list_filter   = ("activo", "campo")
    search_fields = ("campo__nombre", "variedad__nombre")
    autocomplete_fields = ("campo", "variedad")
    ordering      = ("campo__nombre", "variedad__nombre")

@admin.register(SalidaDia)
class SalidaDiaAdmin(admin.ModelAdmin):
    list_display = ("fecha", "campo")
    search_fields = ("campo__nombre", "notas")
    list_filter = ("campo", "fecha")

@admin.register(SalidaItem)
class SalidaItemAdmin(admin.ModelAdmin):
    list_display = ("salida", "variedad", "kg")
    search_fields = ("variedad__nombre",)
    list_filter = ("variedad",)

class ProduccionItemInline(admin.TabularInline):
    model = ProduccionItem
    extra = 0
    autocomplete_fields = ("variedad",)
    fields = ("variedad", "kg")
    readonly_fields = ()
    can_delete = True



@admin.register(ProduccionDia)
class ProduccionDiaAdmin(admin.ModelAdmin):
    date_hierarchy = "fecha"
    list_display   = ("fecha", "campo", "rezaga_kg", "total_general_kg", "empacado_kg", "porcentaje_rezaga_redondeado")
    list_filter    = ("campo",)
    search_fields  = ("campo__nombre",)
    autocomplete_fields = ("campo",)
    inlines        = [ProduccionItemInline]
    readonly_fields = ("total_general_kg", "empacado_kg", "porcentaje_rezaga_mostrable")
    fields = ("fecha", "campo", "rezaga_kg", "notas",
              "total_general_kg", "empacado_kg", "porcentaje_rezaga_mostrable")

    def porcentaje_rezaga_redondeado(self, obj):
        try:
            return round(float(obj.porcentaje_rezaga), 2)
        except Exception:
            return 0.0
    porcentaje_rezaga_redondeado.short_description = "% rezaga"

    def porcentaje_rezaga_mostrable(self, obj):
        return self.porcentaje_rezaga_redondeado(obj)
    porcentaje_rezaga_mostrable.short_description = "% rezaga"

@admin.register(Salida)
class SalidaAdmin(admin.ModelAdmin):
    list_display = ("fecha", "campo", "variedad", "kg", "notas")
    list_filter  = ("campo", "variedad", "fecha")
    search_fields = ("notas",)

from .models import InventarioMovimiento
@admin.register(InventarioMovimiento)
class InventarioMovimientoAdmin(admin.ModelAdmin):
    list_display = ("fecha", "scope", "campo", "variedad", "concepto", "kg", "cs_6oz", "cs_9_8oz", "cs_18oz", "notas")
    list_filter  = ("scope", "concepto", "campo", "variedad")
    search_fields = ("notas",)
    date_hierarchy = "fecha"