# empaques/forms_inventory.py
from django import forms
from .models import InventoryItem, InventoryMovement

class InventoryItemForm(forms.ModelForm):
    class Meta:
        model = InventoryItem
        fields = ["sku", "name", "location", "unit"]
        labels = {
            "sku": "Id",        # ← renombrado
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
