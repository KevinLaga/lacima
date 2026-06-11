# empaques/forms_inventory.py
from django import forms
from django.forms import inlineformset_factory
from .models import InventoryItem, InventoryMovement, Pedimento, PedimentoItem, Remision, RemisionItem


class InventoryItemForm(forms.ModelForm):
    class Meta:
        model = InventoryItem
        fields = ["sku", "name", "location", "unit"]
        labels = {
            "sku": "Id",
            "name": "Artículo",
            "location": "Ubicación",
            "unit": "Unidad",
        }


class InventoryMovementForm(forms.ModelForm):
    class Meta:
        model = InventoryMovement
        fields = ["item", "type", "quantity", "date", "reference", "notes"]
        widgets = {
            "date": forms.DateInput(attrs={"type": "date"}),
        }

    def __init__(self, *args, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.user = user

    def save(self, commit=True):
        obj = super().save(commit=False)
        if self.user and hasattr(obj, "created_by_id") and not obj.created_by_id:
            obj.created_by = self.user
        if commit:
            obj.save()
        return obj


class PedimentoForm(forms.ModelForm):
    class Meta:
        model = Pedimento
        fields = ["empresa", "fecha", "notas"]
        widgets = {
            "fecha": forms.DateInput(attrs={"type": "date"}),
            "notas": forms.Textarea(attrs={"rows": 2}),
        }
        labels = {
            "empresa": "Empresa",
            "fecha": "Fecha",
            "notas": "Notas",
        }


class PedimentoItemForm(forms.ModelForm):
    class Meta:
        model = PedimentoItem
        fields = ["articulo", "cantidad"]
        labels = {
            "articulo": "Artículo",
            "cantidad": "Cantidad",
        }
        widgets = {
            "cantidad": forms.NumberInput(attrs={"min": "0.01", "step": "0.01"}),
        }


PedimentoItemFormSet = inlineformset_factory(
    Pedimento, PedimentoItem,
    form=PedimentoItemForm,
    extra=3,
    min_num=1,
    validate_min=True,
    can_delete=True,
)


class RemisionForm(forms.ModelForm):
    class Meta:
        model = Remision
        fields = ["empresa", "fecha", "notas"]
        widgets = {
            "fecha": forms.DateInput(attrs={"type": "date"}),
            "notas": forms.Textarea(attrs={"rows": 2}),
        }
        labels = {
            "empresa": "Empresa",
            "fecha": "Fecha",
            "notas": "Notas",
        }


class RemisionItemForm(forms.ModelForm):
    class Meta:
        model = RemisionItem
        fields = ["articulo", "cantidad"]
        labels = {
            "articulo": "Artículo",
            "cantidad": "Cantidad",
        }
        widgets = {
            "cantidad": forms.NumberInput(attrs={"min": "0.01", "step": "0.01"}),
        }


RemisionItemFormSet = inlineformset_factory(
    Remision, RemisionItem,
    form=RemisionItemForm,
    extra=3,
    min_num=1,
    validate_min=True,
    can_delete=True,
)
