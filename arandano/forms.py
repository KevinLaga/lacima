# arandano/forms.py
from django import forms
from .models import Campo, Variedad, CampoVariedad, ProduccionDia, ProduccionItem, SalidaDia, SalidaItem, Salida, DestinoSalida
from django.forms import formset_factory

class SalidaForm(forms.ModelForm):
    fecha = forms.DateField(widget=forms.DateInput(attrs={"type": "date"}))

    class Meta:
        model = Salida
        fields = ["fecha", "campo", "variedad", "kg", "notas"]
        widgets = {
            "kg": forms.NumberInput(attrs={"step": "0.00001", "min": "0"}),
        }

    def __init__(self, *args, **kwargs):
        campo_sel = kwargs.pop("campo_sel", None)
        super().__init__(*args, **kwargs)
        self.fields["campo"].queryset = Campo.objects.filter(activo=True).order_by("nombre")
        self.fields["variedad"].required = False
        if campo_sel:
            # Si llega un campo preseleccionado, filtra sus variedades
            vids = CampoVariedad.objects.filter(campo=campo_sel, activo=True, variedad__activo=True)\
                                        .values_list("variedad_id", flat=True)
            self.fields["variedad"].queryset = Variedad.objects.filter(pk__in=vids).order_by("nombre")
        else:
            self.fields["variedad"].queryset = Variedad.objects.filter(activo=True).order_by("nombre")
class SalidaDiaForm(forms.ModelForm):
    class Meta:
        model = SalidaDia
        fields = ["fecha", "campo", "destino", "destino_detalle", "notas"]

    def clean(self):
        cleaned = super().clean()
        destino = cleaned.get("destino")
        detalle = (cleaned.get("destino_detalle") or "").strip()

        # Si viene vacío por algún motivo, fuerza default para evitar "obligatorio"
        if not destino:
            destino = DestinoSalida.EMPAQUE
            cleaned["destino"] = destino

        if destino == DestinoSalida.OTRO and not detalle:
            self.add_error("destino_detalle", "Especifica el destino.")

        if destino == DestinoSalida.EMPAQUE:
            cleaned["destino_detalle"] = ""

        return cleaned


class SalidaItemForm(forms.Form):
    variedad = forms.ModelChoiceField(queryset=Variedad.objects.none())
    kg = forms.DecimalField(max_digits=12, decimal_places=5, min_value=0, required=False)

    cs_6oz   = forms.IntegerField(required=False, min_value=0, initial=0)
    cs_9_8oz = forms.IntegerField(required=False, min_value=0, initial=0)
    cs_18oz  = forms.IntegerField(required=False, min_value=0, initial=0)

    def __init__(self, *args, **kwargs):
        campo = kwargs.pop("campo", None)
        super().__init__(*args, **kwargs)
        qs = Variedad.objects.none()
        if campo:
            qs = Variedad.objects.filter(
                id__in=CampoVariedad.objects.filter(campo=campo, activo=True).values_list("variedad_id", flat=True),
                activo=True
            ).order_by("nombre")
        self.fields["variedad"].queryset = qs

SalidaItemFormSet = formset_factory(SalidaItemForm, extra=0, can_delete=False)

class ProduccionDiaForm(forms.ModelForm):
    rezaga_kg = forms.DecimalField(required=False, initial=0)
    class Meta:
        model = ProduccionDia
        fields = ("fecha", "campo", "rezaga_kg", "notas")
        widgets = {
            # calendario nativo (igual que en embarques si usas input date)
            "fecha": forms.DateInput(
                format="%Y-%m-%d",
                attrs={
                    "type": "date",          # fuerza el calendario HTML5
                    "class": "date-input",   # por si en embarques usas CSS/JS a esta clase
                }
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # acepta el formato del input date
        self.fields["fecha"].input_formats = ["%Y-%m-%d"]


class ProduccionItemForm(forms.Form):
    variedad = forms.ModelChoiceField(queryset=Variedad.objects.none(), required=True)
    kg = forms.DecimalField(required=False, min_value=0, decimal_places=5, max_digits=10)
    cs_6oz   = forms.IntegerField(required=False, min_value=0, initial=0)
    cs_9_8oz = forms.IntegerField(required=False, min_value=0, initial=0)
    cs_18oz  = forms.IntegerField(required=False, min_value=0, initial=0)

from django.forms import formset_factory
ProduccionItemFormSet = formset_factory(ProduccionItemForm, extra=0, can_delete=False)

