from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal

from django.db import models
from django.db.models import Sum
from django.utils import timezone


# ─────────────────────────── Catálogos ───────────────────────────

EMPRESA_CHOICES = [
    ("LA CIMA",    "La Cima"),
    ("EMPAQUE N1", "Empaque N1"),
    ("RC",         "RC"),
]

BANCO_CHOICES = [
    ("BBVA",    "BBVA"),
    ("BAJIO",   "Bajío"),
    ("HSBC",    "HSBC"),
    ("ELEVATE", "Elevate Export Finance"),
]

TIPO_CREDITO_CHOICES = [
    ("CTA_CORRIENTE",  "Cuenta corriente"),
    ("SIMPLE",         "Crédito simple"),
    ("CC_GARANTIA_H",  "CC garantía hipotecaria"),
    ("OTRO",           "Otro (especificar)"),
]

MONEDA_CHOICES = [
    ("MXN", "Pesos MXN"),
    ("USD", "Dólares USD"),
]

FRECUENCIA_CHOICES = [
    ("MENSUAL",    "Mensual"),
    ("BIMESTRAL",  "Cada 2 meses"),
    ("TRIMESTRAL", "Cada 3 meses"),
    ("SEMESTRAL",  "Cada 6 meses"),
    ("ANUAL",      "Anual"),
]

FRECUENCIA_MESES = {
    "MENSUAL": 1, "BIMESTRAL": 2, "TRIMESTRAL": 3, "SEMESTRAL": 6, "ANUAL": 12,
}

MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun",
            "jul", "ago", "sep", "oct", "nov", "dic"]


def sumar_meses(fecha: date, meses: int) -> date:
    """
    Suma meses conservando el día. Si el día no existe en el mes destino
    (ej. 31 de enero + 1 mes), cae al último día de ese mes.
    """
    total = fecha.month - 1 + meses
    anio = fecha.year + total // 12
    mes = total % 12 + 1
    # Último día del mes destino
    if mes == 12:
        ultimo = 31
    else:
        ultimo = (date(anio, mes + 1, 1) - timedelta(days=1)).day
    return date(anio, mes, min(fecha.day, ultimo))

SIMBOLO_MONEDA = {"MXN": "$", "USD": "US$"}


class Garantia(models.Model):
    """
    Terreno o campo dado en garantía al banco.

    Es un catálogo para poder cargar la lista de campos sin tocar código:
    se administran desde el admin de Django.
    """
    nombre      = models.CharField("Nombre del campo / terreno", max_length=120, unique=True)
    descripcion = models.CharField("Descripción", max_length=255, blank=True)
    precio      = models.DecimalField("Precio", max_digits=14, decimal_places=2,
                                      null=True, blank=True)
    activo      = models.BooleanField("Activo", default=True)

    class Meta:
        verbose_name = "Garantía"
        verbose_name_plural = "Garantías"
        ordering = ["nombre"]

    def __str__(self):
        return self.nombre

    @property
    def precio_fmt(self):
        if self.precio is None:
            return "—"
        return f"${self.precio:,.2f}"

    @property
    def en_uso(self) -> bool:
        """True si algún crédito la tiene como garantía (no se puede eliminar)."""
        return self.creditos.exists()


# ─────────────────────────── Crédito ───────────────────────────

class Credito(models.Model):
    empresa = models.CharField("Empresa", max_length=20, choices=EMPRESA_CHOICES)
    banco   = models.CharField("Banco",   max_length=20, choices=BANCO_CHOICES)

    tipo_credito = models.CharField("Tipo de crédito", max_length=20, choices=TIPO_CREDITO_CHOICES)
    tipo_otro    = models.CharField(
        "Especificar tipo", max_length=120, blank=True,
        help_text="Sólo si el tipo de crédito es 'Otro'.",
    )

    garantia = models.ForeignKey(
        Garantia, verbose_name="Garantía", on_delete=models.PROTECT,
        null=True, blank=True, related_name="creditos",
    )

    moneda = models.CharField("Moneda", max_length=3, choices=MONEDA_CHOICES, default="MXN")
    tasa   = models.DecimalField("Tasa (%)", max_digits=6, decimal_places=3,
                                 null=True, blank=True)

    plazo_meses = models.PositiveIntegerField(
        "Plazo (meses)", null=True, blank=True,
        help_text="Duración del crédito en meses.",
    )

    cantidad_pagos = models.PositiveIntegerField(
        "Cantidad de pagos", null=True, blank=True,
        help_text="En cuántas exhibiciones se liquida el monto del crédito.",
    )
    frecuencia_pagos = models.CharField(
        "Tiempo de pagos", max_length=12, choices=FRECUENCIA_CHOICES, blank=True,
        help_text="Cada cuánto se paga.",
    )

    monto = models.DecimalField("Monto del crédito", max_digits=14, decimal_places=2)

    fecha_contratacion = models.DateField(
        "Fecha de crédito", null=True, blank=True,
        help_text="Fecha en que se contrató el crédito.",
    )
    fecha_disposicion = models.DateField("Fecha de disposición")
    fecha_vencimiento = models.DateField(
        "Fecha de vencimiento",
        help_text="Fecha en que vence el crédito.",
    )

    referencia = models.CharField("Núm. de crédito / referencia", max_length=60, blank=True)
    notas      = models.TextField("Notas", blank=True)

    creado_en      = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Crédito"
        verbose_name_plural = "Créditos"
        ordering = ["fecha_vencimiento", "id"]

    def __str__(self):
        return f"{self.get_empresa_display()} · {self.get_banco_display()} · {self.monto_fmt}"

    # ── Etiquetas ──

    @property
    def tipo_credito_label(self):
        if self.tipo_credito == "OTRO":
            return self.tipo_otro or "Otro"
        return self.get_tipo_credito_display()

    @property
    def simbolo(self):
        return SIMBOLO_MONEDA.get(self.moneda, "$")

    @property
    def monto_fmt(self):
        return f"{self.simbolo}{self.monto:,.2f}"

    @property
    def anio_vencimiento(self):
        return self.fecha_vencimiento.year if self.fecha_vencimiento else None

    # ── Intereses ──

    @property
    def interes(self) -> Decimal:
        """
        Interés a pagar: la tasa aplicada como porcentaje sobre el monto.

        Es un porcentaje plano sobre el capital (ej. tasa 4.15 sobre un monto
        de 100,000 = 4,150). No se prorratea por tiempo ni se capitaliza.
        """
        if not self.monto or self.tasa is None:
            return Decimal("0.00")
        bruto = Decimal(self.monto) * Decimal(self.tasa) / Decimal("100")
        return bruto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    @property
    def total_a_pagar(self) -> Decimal:
        """
        Lo que hay que cubrir para liquidar: SOLO el monto del crédito.

        Los intereses de la tasa se manejan aparte (la tasa cambia con el tiempo
        y en créditos de varios años no se puede fijar de antemano), así que
        'interes' queda como dato informativo y no entra aquí.
        """
        return Decimal(self.monto or 0)

    # ── Abonos y saldo ──

    @property
    def total_abonado(self) -> Decimal:
        total = self.abonos.aggregate(t=Sum("monto"))["t"]
        return Decimal(total or 0)

    @property
    def saldo(self) -> Decimal:
        """Lo que falta por abonar del monto del crédito."""
        return self.total_a_pagar - self.total_abonado

    @property
    def interes_fmt(self):
        return f"{self.simbolo}{self.interes:,.2f}"

    @property
    def total_a_pagar_fmt(self):
        return f"{self.simbolo}{self.total_a_pagar:,.2f}"

    @property
    def saldo_fmt(self):
        return f"{self.simbolo}{self.saldo:,.2f}"

    @property
    def abonado_fmt(self):
        return f"{self.simbolo}{self.total_abonado:,.2f}"

    @property
    def porcentaje_pagado(self) -> float:
        total = self.total_a_pagar
        if not total:
            return 0.0
        pct = float(self.total_abonado) / float(total) * 100.0
        return max(0.0, min(100.0, pct))

    @property
    def liquidado(self) -> bool:
        return self.saldo <= Decimal("0.005")

    # ── Vencimiento ──

    @property
    def dias_para_vencer(self):
        """Días que faltan. Negativo si ya venció. None si no hay fecha."""
        if not self.fecha_vencimiento:
            return None
        return (self.fecha_vencimiento - timezone.localdate()).days

    @property
    def estado(self) -> str:
        """liquidado | vencido | urgente | proximo | vigente"""
        if self.liquidado:
            return "liquidado"
        d = self.dias_para_vencer
        if d is None:
            return "vigente"
        if d < 0:
            return "vencido"
        if d <= 30:
            return "urgente"
        if d <= 90:
            return "proximo"
        return "vigente"

    @property
    def estado_label(self) -> str:
        return {
            "liquidado": "Liquidado",
            "vencido":   "Vencido",
            "urgente":   "Por vencer",
            "proximo":   "Próximo",
            "vigente":   "Vigente",
        }.get(self.estado, "Vigente")

    # ── Plan de pagos ──

    @property
    def tiene_plan(self) -> bool:
        return bool(self.cantidad_pagos and self.frecuencia_pagos and self.monto)

    @property
    def monto_por_pago(self) -> Decimal:
        """Monto del crédito repartido entre la cantidad de pagos."""
        if not self.tiene_plan:
            return Decimal("0.00")
        bruto = Decimal(self.monto) / Decimal(self.cantidad_pagos)
        return bruto.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    @property
    def monto_por_pago_fmt(self):
        return f"{self.simbolo}{self.monto_por_pago:,.2f}"

    @property
    def frecuencia_label(self):
        return self.get_frecuencia_pagos_display() if self.frecuencia_pagos else "—"

    def plan_pagos(self):
        """
        Calendario de pagos calculado a partir del monto, la cantidad de pagos y
        la frecuencia. El primer pago cae un periodo después de la disposición.

        Devuelve una lista de dicts:
          num, fecha, monto, acumulado, pagado, anio, fecha_texto
        Un pago se marca 'pagado' cuando los abonos acumulados alcanzan a
        cubrirlo (los abonos se aplican en orden contra el calendario).
        """
        if not self.tiene_plan or not self.fecha_disposicion:
            return []

        paso = FRECUENCIA_MESES.get(self.frecuencia_pagos, 1)
        n = int(self.cantidad_pagos)
        cuota = self.monto_por_pago
        abonado = self.total_abonado

        filas = []
        acumulado = Decimal("0.00")
        for i in range(1, n + 1):
            # El último pago absorbe el redondeo para que la suma dé el monto exacto
            if i == n:
                importe = Decimal(self.monto) - acumulado
            else:
                importe = cuota
            acumulado += importe

            fecha = sumar_meses(self.fecha_disposicion, paso * i)
            filas.append({
                "num": i,
                "fecha": fecha,
                "anio": fecha.year,
                "fecha_texto": f"{fecha.day} {MESES_ES[fecha.month - 1]}-{fecha.year}",
                "monto": importe,
                "monto_fmt": f"{self.simbolo}{importe:,.2f}",
                "acumulado": acumulado,
                "pagado": abonado >= acumulado - Decimal("0.005"),
            })
        return filas

    def pagos_pendientes(self):
        """Sólo los pagos que todavía no quedan cubiertos por los abonos."""
        return [p for p in self.plan_pagos() if not p["pagado"]]

    @property
    def proximo_pago(self):
        pendientes = self.pagos_pendientes()
        return pendientes[0] if pendientes else None

    @property
    def vencimiento_texto(self) -> str:
        """Texto legible de cuánto falta o cuánto lleva vencido."""
        d = self.dias_para_vencer
        if d is None:
            return "—"
        if self.liquidado:
            return "Liquidado"
        if d < 0:
            n = abs(d)
            return f"Venció hace {n} día{'s' if n != 1 else ''}"
        if d == 0:
            return "Vence hoy"
        return f"Faltan {d} día{'s' if d != 1 else ''}"


class Abono(models.Model):
    credito = models.ForeignKey(
        Credito, verbose_name="Crédito", on_delete=models.CASCADE, related_name="abonos",
    )
    fecha      = models.DateField("Fecha del abono", default=date.today)
    monto      = models.DecimalField("Monto abonado", max_digits=14, decimal_places=2)
    referencia = models.CharField("Referencia", max_length=60, blank=True)
    nota       = models.CharField("Nota", max_length=255, blank=True)

    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Abono"
        verbose_name_plural = "Abonos"
        ordering = ["-fecha", "-id"]

    def __str__(self):
        return f"{self.fecha} · {self.monto}"

    @property
    def monto_fmt(self):
        return f"{self.credito.simbolo}{self.monto:,.2f}"
