# empaques/forms.py
from django import forms
from django.forms import ModelForm
from django.forms.models import BaseInlineFormSet
from django.core.exceptions import ValidationError
from .models import Shipment, ShipmentItem, Presentation

class BaseShipmentItemFormSet(BaseInlineFormSet):
    def clean(self):
        super().clean()
        for form in self.forms:
            if getattr(form, 'cleaned_data', None) is None or form.cleaned_data.get('DELETE', False):
                continue

            filled = any([
                form.cleaned_data.get('tarima'),
                form.cleaned_data.get('type'),
                form.cleaned_data.get('quantity'),
            ])
            if filled:
                for field in ['tarima', 'type', 'size', 'quantity']:
                    if not form.cleaned_data.get(field):
                        form.add_error(field, "Este campo es requerido.")
            else:
                form.cleaned_data['DELETE'] = True


class ShipmentForm(forms.ModelForm):
    date = forms.DateField(
        label="Fecha",
        widget=forms.DateInput(attrs={'type': 'date'})
    )
    tarimas_peco = forms.IntegerField(
        label="Tarimas PECO",
        min_value=0,
        required=False,
        widget=forms.NumberInput(attrs={'placeholder': 'Ej: 3'})
    )
    departure_time = forms.TimeField(
        label="Horario de salida",
        widget=forms.TimeInput(attrs={'type': 'time'})
    )
    class Meta:
        model = Shipment
        fields = [
            'tracking_number','date','carrier','tractor_plates','box_plates',
            'driver','departure_time','box','box_conditions','box_free_of_odors',
            'ryan','seal_1','seal_2','seal_3','seal_4','chismografo',
            'delivery_signature','driver_signature','invoice_number','tarimas_peco',
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

    # 👇 Evitamos query en import: arrancamos con .none() y lo rellenamos en __init__
    type = forms.ModelChoiceField(
        queryset=Presentation.objects.none(),
        label="Tipo"
    )

    cliente = forms.ChoiceField(
        choices=CLIENTE_CHOICES,
        required=False,
        label="Cliente"
    )

    temperatura = forms.DecimalField(
        required=False,
        label="Temperatura (°F)",
        max_digits=5,
        decimal_places=1,
        widget=forms.NumberInput(attrs={'step': '0.1', 'placeholder': 'Ej: 36.2'})
    )

    class Meta:
        model = ShipmentItem
        fields = ['type', 'size', 'quantity', 'tarima', 'cliente', 'temperatura']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ❗ Aquí ya es seguro asignar el queryset
        self.fields['type'].queryset = Presentation.objects.order_by('name')

    def clean(self):
        cleaned = super().clean()
        # mantén tu lógica si necesitas setear instance.presentation, etc.
        self.instance.presentation = cleaned.get('type')
        return cleaned




