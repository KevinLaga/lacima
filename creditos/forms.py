from django import forms

from .models import FRECUENCIA_CHOICES, Abono, Credito, Garantia


class CreditoForm(forms.ModelForm):
    class Meta:
        model = Credito
        fields = [
            "empresa", "banco", "tipo_credito", "tipo_otro", "garantia",
            "moneda", "tasa", "plazo_meses", "cantidad_pagos", "frecuencia_pagos",
            "monto",
            "fecha_contratacion", "fecha_disposicion", "fecha_vencimiento",
            "referencia", "notas",
        ]
        widgets = {
            "fecha_contratacion": forms.DateInput(attrs={"type": "date"}),
            "fecha_disposicion": forms.DateInput(attrs={"type": "date"}),
            "fecha_vencimiento": forms.DateInput(attrs={"type": "date"}),
            "tasa":        forms.NumberInput(attrs={"step": "0.001", "placeholder": "Ej: 12.500"}),
            "monto":       forms.NumberInput(attrs={"step": "0.01",  "placeholder": "Ej: 1500000.00"}),
            "plazo_meses": forms.NumberInput(attrs={"placeholder": "Ej: 12"}),
            "cantidad_pagos": forms.NumberInput(attrs={"placeholder": "Ej: 6", "min": "1"}),
            "tipo_otro":   forms.TextInput(attrs={"placeholder": "Sólo si elegiste 'Otro'"}),
            "referencia":  forms.TextInput(attrs={"placeholder": "Núm. de crédito (opcional)"}),
            "notas":       forms.Textarea(attrs={"rows": 3}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["garantia"].queryset = Garantia.objects.filter(activo=True).order_by("nombre")
        self.fields["garantia"].empty_label = "— Sin garantía —"
        self.fields["frecuencia_pagos"].widget.choices = (
            [("", "— Selecciona —")] + list(FRECUENCIA_CHOICES)
        )

    def clean(self):
        cleaned = super().clean()

        if cleaned.get("tipo_credito") == "OTRO" and not (cleaned.get("tipo_otro") or "").strip():
            self.add_error("tipo_otro", "Especifica el tipo de crédito.")

        monto = cleaned.get("monto")
        if monto is not None and monto <= 0:
            self.add_error("monto", "El monto debe ser mayor a cero.")

        # Cantidad y frecuencia van juntas: con una sola no se puede armar el calendario
        cant = cleaned.get("cantidad_pagos")
        frec = cleaned.get("frecuencia_pagos")
        if cant is not None and cant < 1:
            self.add_error("cantidad_pagos", "Debe ser al menos 1 pago.")
        if cant and not frec:
            self.add_error("frecuencia_pagos",
                           "Indica cada cuánto se paga para poder calcular las fechas.")
        if frec and not cant:
            self.add_error("cantidad_pagos",
                           "Indica en cuántos pagos se liquida para poder calcular las fechas.")

        f_contrat = cleaned.get("fecha_contratacion")
        f_disp = cleaned.get("fecha_disposicion")
        f_venc = cleaned.get("fecha_vencimiento")
        if f_disp and f_venc and f_venc < f_disp:
            self.add_error("fecha_vencimiento",
                           "El vencimiento no puede ser anterior a la disposición.")
        if f_contrat and f_disp and f_disp < f_contrat:
            self.add_error("fecha_disposicion",
                           "La disposición no puede ser anterior a la fecha de crédito.")
        return cleaned


class AbonoForm(forms.ModelForm):
    class Meta:
        model = Abono
        fields = ["fecha", "monto", "referencia", "nota"]
        widgets = {
            "fecha":      forms.DateInput(attrs={"type": "date"}),
            "monto":      forms.NumberInput(attrs={"step": "0.01", "placeholder": "0.00"}),
            "referencia": forms.TextInput(attrs={"placeholder": "Folio / transferencia"}),
            "nota":       forms.TextInput(attrs={"placeholder": "Opcional"}),
        }

    def __init__(self, *args, credito=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.credito = credito

    def clean_monto(self):
        monto = self.cleaned_data["monto"]
        if monto is None or monto <= 0:
            raise forms.ValidationError("El abono debe ser mayor a cero.")

        if self.credito is not None:
            # Saldo disponible, sin contar este abono si se está editando
            saldo = self.credito.saldo
            if self.instance.pk:
                saldo += self.instance.monto
            if monto > saldo:
                raise forms.ValidationError(
                    f"El abono excede el saldo pendiente "
                    f"({self.credito.simbolo}{saldo:,.2f})."
                )
        return monto


class GarantiaForm(forms.ModelForm):
    class Meta:
        model = Garantia
        fields = ["nombre", "descripcion", "precio"]
        widgets = {
            "nombre":      forms.TextInput(attrs={"placeholder": "Ej: campo 21 misión"}),
            "descripcion": forms.TextInput(attrs={"placeholder": "Superficie, ubicación, etc."}),
            "precio":      forms.NumberInput(attrs={"step": "0.01", "placeholder": "0.00"}),
        }

    def clean_precio(self):
        precio = self.cleaned_data.get("precio")
        if precio is not None and precio < 0:
            raise forms.ValidationError("El precio no puede ser negativo.")
        return precio
