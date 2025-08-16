from django import forms
from django.forms import ModelForm
from django.forms.models import BaseInlineFormSet
from django.core.exceptions import ValidationError
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import re
from .models import Shipment, ShipmentItem, Presentation


class BaseShipmentItemFormSet(BaseInlineFormSet):
    def clean(self):
        super().clean()
        for form in self.forms:
            if getattr(form, 'cleaned_data', None) is None or form.cleaned_data.get('DELETE', False):
                continue

            # Solo tomamos en cuenta tarima, type y quantity para considerar si está lleno
            filled = any([
                form.cleaned_data.get('tarima'),
                form.cleaned_data.get('type'),
                form.cleaned_data.get('quantity')
            ])
            if filled:
                # Si hay alguno de estos campos, ahora sí exigimos todos
                for field in ['tarima', 'type', 'size', 'quantity']:
                    if not form.cleaned_data.get(field):
                        form.add_error(field, "Este campo es requerido.")
            else:
                # Si está vacío (ignorar size por default), lo marcamos para borrar
                form.cleaned_data['DELETE'] = True


class ShipmentForm(forms.ModelForm):
    date = forms.DateField(
        label="Fecha",
        widget=forms.DateInput(attrs={'type': 'date'})
    )
    departure_time = forms.TimeField(
        label="Horario de salida",
        widget=forms.TimeInput(attrs={'type': 'time'})
    )

    class Meta:
        model = Shipment
        fields = [
            'tracking_number',   # Número de orden
            'date',              # Fecha
            'carrier',           # Transportista
            'tractor_plates',    # Placas tractor
            'box_plates',        # Placas caja
            'driver',            # Operador
            'departure_time',    # Horario de salida
            'box',               # Caja
            'box_conditions',    # Condiciones de la caja
            'box_free_of_odors', # Caja libre de olores
            'ryan',              # Ryan
            'seal_1', 'seal_2', 'seal_3', 'seal_4', # Sellos
            'chismografo',
            'delivery_signature',# Firma de entrega
            'driver_signature',  # Firma de operador
            'invoice_number',    # Número de factura
        ]


CLIENTE_CHOICES = [
    ('La Cima Produce', 'La Cima Produce'),
    ('RC Organics', 'RC Organics'),
    ('GH Farms', 'GH Farms'),
    ('Gourmet Baja Farms', 'Gourmet Baja Farms'),
    ('GBF Farms', 'GBF Farms'),
]


class ShipmentItemForm(ModelForm):
    tarima = forms.IntegerField(label="Tarima", min_value=1, max_value=26)

    # dropdown de Presentation
    type = forms.ModelChoiceField(
        queryset=Presentation.objects.order_by('name'),
        label="Tipo"
    )

    cliente = forms.ChoiceField(
        choices=CLIENTE_CHOICES,
        required=False, 
        label="Cliente"
    )

    temperatura = forms.CharField(
        required=False,
        label="Temperatura",
        widget=forms.TextInput(attrs={'placeholder': 'Ej: 5°C'})
    )

    class Meta:
        model = ShipmentItem
        fields = ['type', 'size', 'quantity', 'tarima', 'cliente', 'temperatura']

    def clean_temperatura(self):
        val = self.cleaned_data.get('temperatura')
        if val in (None, ''):
            return None

        # Si ya es número
        if isinstance(val, (int, float, Decimal)):
            return Decimal(str(val)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)

        s = str(val).strip()
        s = s.replace(',', '.')
        m = re.search(r'-?\d+(?:\.\d+)?', s)
        if not m:
            raise ValidationError("Ingresa una temperatura válida (ej. 36.2°F).")

        try:
            num = Decimal(m.group(0)).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            raise ValidationError("Ingresa una temperatura válida (ej. 36.2°F).")

        return num

    def clean(self):
        cleaned = super().clean()
        # "type" es el objeto Presentation
        self.instance.presentation = cleaned.get('type')
        return cleaned




