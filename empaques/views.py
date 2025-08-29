import os
import csv
from datetime import date
from collections import defaultdict

from django.conf import settings
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.forms import inlineformset_factory
from django.http import HttpResponse, HttpResponseForbidden
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect


from .models import Shipment, ShipmentItem
from .forms import (
    ShipmentForm,
    ShipmentItemForm,
    BaseShipmentItemFormSet,
)
from django.contrib.auth.decorators import login_required, permission_required
def es_capturista(user):
    return user.is_authenticated and user.groups.filter(name="capturista").exists()

@login_required
def post_login_redirect(request):
    """
    Redirige según el rol:
    - capturista -> nuevo embarque
    - demás -> lista de embarques
    """
    if es_capturista(request.user):
        return redirect("shipment_create")
    return redirect("shipment_list")

from django.utils.crypto import get_random_string
from django.db import transaction
from django.forms import inlineformset_factory
from django.contrib.auth.decorators import login_required, permission_required
@login_required
@permission_required('empaques.add_shipment', raise_exception=True)
def shipment_create(request):
    ItemFormSet = inlineformset_factory(
        Shipment,
        ShipmentItem,
        form=ShipmentItemForm,
        formset=BaseShipmentItemFormSet,
        extra=26,
        can_delete=True,
    )

    # Inicializa la lista de tokens usados en sesión
    used_tokens = request.session.get('used_form_tokens', [])
    if not isinstance(used_tokens, list):
        used_tokens = []
    request.session['used_form_tokens'] = used_tokens

    if request.method == 'POST':
        token = request.POST.get('form_token')
        used = set(request.session.get('used_form_tokens', []))

        # Si no hay token o ya se usó, no duplica
        if (not token) or (token in used):
            return redirect('shipment_list')

        form = ShipmentForm(request.POST)
        formset = ItemFormSet(request.POST)

        if form.is_valid() and formset.is_valid():
            # Marcamos token como usado ANTES de guardar (bloquea multi-clics)
            used.add(token)
            request.session['used_form_tokens'] = list(used)
            request.session.modified = True

            with transaction.atomic():
                shipment = form.save()
                formset.instance = shipment
                formset.save()

            return redirect('shipment_list')

        # Si hay errores, generamos un nuevo token
        new_token = get_random_string(32)
        request.session['current_form_token'] = new_token
        request.session.modified = True
        return render(request, 'empaques/shipment_form.html', {
            'form': form,
            'formset': formset,
            'form_token': new_token,
        })

    # GET: token nuevo
    token = get_random_string(32)
    request.session['current_form_token'] = token
    # Limpieza de tokens antiguos
    if len(used_tokens) > 100:
        request.session['used_form_tokens'] = used_tokens[-50:]
    request.session.modified = True

    form = ShipmentForm()
    formset = ItemFormSet()
    return render(request, 'empaques/shipment_form.html', {
        'form': form,
        'formset': formset,
        'form_token': token,
    })
from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.http import HttpResponse
from datetime import date
from collections import defaultdict

from .models import Shipment, ShipmentItem


@login_required
def shipment_list(request):
    from datetime import date, timedelta
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    # --- Lista de embarques para la tabla ---
    shipments = Shipment.objects.order_by('-date', '-id')

    # --- Permiso para descargar/exportar ---
    can_download = request.user.has_perm('empaques.can_download_reports')

    # --- Parámetros de periodo ---
    try:
        year = int(request.GET.get('year') or date.today().year)
    except ValueError:
        year = date.today().year
    try:
        month = int(request.GET.get('month') or date.today().month)
    except ValueError:
        month = date.today().month

    # --------------------------
    # Helper: genera XLSX común
    # --------------------------
    

    def build_summary_xlsx(title_text, subtitle_text, embarques_qs):
        items = ShipmentItem.objects.filter(shipment__in=embarques_qs).select_related('presentation')

        # Agregados por presentación/tamaño
        presentaciones_info = defaultdict(lambda: {'cajas': 0, 'dinero': 0.0})
        for it in items:
            k = (it.presentation.name, it.size)
            presentaciones_info[k]['cajas'] += it.quantity
            presentaciones_info[k]['dinero'] += it.quantity * float(it.presentation.price)

        total_cajas = sum(i.quantity for i in items)
        total_eq_11lbs = sum(i.quantity * float(i.presentation.conversion_factor) for i in items)
        total_dinero = sum(i.quantity * float(i.presentation.price) for i in items)

        wb = Workbook()
        ws = wb.active
        ws.title = "Resumen"

        # Estilos
        title_font = Font(name="Calibri", size=16, bold=True, color="3C78D8")
        h_font = Font(name="Calibri", size=12, bold=True)
        th_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        th_fill = PatternFill("solid", fgColor="225577")
        border = Border(
            left=Side(style='thin', color='AAAAAA'),
            right=Side(style='thin', color='AAAAAA'),
            top=Side(style='thin', color='AAAAAA'),
            bottom=Side(style='thin', color='AAAAAA'),
        )

        r = 1
        ws.cell(row=r, column=1, value=title_text).font = title_font
        r += 1
        if subtitle_text:
            ws.cell(row=r, column=1, value=subtitle_text)
            r += 2

        # Presentaciones
        ws.cell(row=r, column=1, value="Presentaciones utilizadas").font = h_font
        r += 1
        headers_pres = ["Presentación", "Tamaño", "Total cajas", "Total dinero"]
        for c, txt in enumerate(headers_pres, start=1):
            cell = ws.cell(row=r, column=c, value=txt)
            cell.font = th_font
            cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        r += 1

        if presentaciones_info:
            for (n_pres, sz), info in sorted(presentaciones_info.items()):
                ws.cell(row=r, column=1, value=n_pres)
                ws.cell(row=r, column=2, value=sz)
                ws.cell(row=r, column=3, value=info['cajas'])
                ws.cell(row=r, column=4, value=round(info['dinero'], 2))
                for c in range(1, 5):
                    ws.cell(row=r, column=c).border = border
                    ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
                r += 1
        else:
            ws.cell(row=r, column=1, value="(Sin datos)")
            r += 1

        r += 1
        # Totales
        ws.cell(row=r, column=1, value="Número total de cajas:").font = h_font
        ws.cell(row=r, column=2, value=total_cajas)
        r += 1
        ws.cell(row=r, column=1, value="Total equivalente en 11 lbs:").font = h_font
        ws.cell(row=r, column=2, value=round(total_eq_11lbs, 2))
        r += 1
        ws.cell(row=r, column=1, value="Total de dinero:").font = h_font
        ws.cell(row=r, column=2, value=round(total_dinero, 2))
        r += 2

        # Detalle
        ws.cell(row=r, column=1, value="Detalle de embarques").font = h_font
        r += 1
        headers_det = ["Fecha", "N# Embarque", "N# Factura", "Presentación", "Tamaño", "Cantidad", "Importe"]
        for c, txt in enumerate(headers_det, start=1):
            cell = ws.cell(row=r, column=c, value=txt)
            cell.font = th_font
            cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        r += 1

        for sh in embarques_qs.order_by('date', 'tracking_number'):
            for it in sh.items.all():
                importe = it.quantity * float(it.presentation.price)
                ws.cell(row=r, column=1, value=sh.date.strftime("%d/%m/%Y"))
                ws.cell(row=r, column=2, value=sh.tracking_number)
                ws.cell(row=r, column=3, value=sh.invoice_number)
                ws.cell(row=r, column=4, value=it.presentation.name)
                ws.cell(row=r, column=5, value=it.size)
                ws.cell(row=r, column=6, value=it.quantity)
                ws.cell(row=r, column=7, value=round(importe, 2))
                for c in range(1, 7 + 1):
                    ws.cell(row=r, column=c).border = border
                    ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
                r += 1

        # Anchos
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 26
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 20.29
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['L'].width = 20.29

        return wb

    # Solo aceptamos descargas si tiene permiso
    descargar = request.GET.get('descargar') if can_download else None

    # --------------------------
    # Descarga SEMANAL (XLSX)
    # --------------------------
    if descargar == 'semana':
        if not request.user.has_perm('empaques.can_download_reports'):
            return HttpResponse("No tienes permiso para descargar reportes.", status=403)

        iso_week = request.GET.get('iso_week', '')
        week = None
        if iso_week:
            try:
                y_str, w_str = iso_week.split('-W')
                year = int(y_str)
                week = int(w_str)
            except Exception:
                week = None

        if week is None:
            try:
                week = int(request.GET.get('week') or date.today().isocalendar().week)
            except ValueError:
                week = date.today().isocalendar().week

        try:
            week_start = date.fromisocalendar(year, week, 1)
            week_end   = date.fromisocalendar(year, week, 7)
        except Exception:
            today = date.today()
            iso = today.isocalendar()
            year, week = iso.year, iso.week
            week_start = date.fromisocalendar(year, week, 1)
            week_end   = date.fromisocalendar(year, week, 7)

        embarques = Shipment.objects.filter(date__range=(week_start, week_end)).order_by('date', 'tracking_number')
        total_embarques = embarques.count()

        items = ShipmentItem.objects.filter(shipment__in=embarques).select_related('presentation')

        presentaciones_info = defaultdict(lambda: {'cajas': 0, 'dinero': 0.0})
        for item in items:
            key = (item.presentation.name, item.size)
            presentaciones_info[key]['cajas'] += item.quantity
            presentaciones_info[key]['dinero'] += item.quantity * float(item.presentation.price)

        total_cajas = sum(i.quantity for i in items)
        total_eq_11lbs = sum(i.quantity * float(i.presentation.conversion_factor) for i in items)
        total_dinero = sum(i.quantity * float(i.presentation.price) for i in items)

        wb = Workbook()
        ws = wb.active
        ws.title = "Resumen semanal"

        title_font = Font(name="Calibri", size=16, bold=True, color="3C78D8")
        h_font = Font(name="Calibri", size=12, bold=True)
        th_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        th_fill = PatternFill("solid", fgColor="225577")
        border = Border(
            left=Side(style='thin', color='AAAAAA'),
            right=Side(style='thin', color='AAAAAA'),
            top=Side(style='thin', color='AAAAAA'),
            bottom=Side(style='thin', color='AAAAAA'),
        )

        r = 1
        ws.cell(row=r, column=1, value=f"Resumen de Embarques – Semana ISO {week} / {year}").font = title_font
        r += 1
        ws.cell(row=r, column=1, value=f"Rango: {week_start.strftime('%d/%m/%Y')} – {week_end.strftime('%d/%m/%Y')}")
        r += 1
        ws.cell(row=r, column=1, value=f"Total de embarques: {total_embarques}")
        r += 2

        ws.cell(row=r, column=1, value="Presentaciones utilizadas").font = h_font
        r += 1
        headers_pres = ["Presentación", "Tamaño", "Total cajas", "Total dinero"]
        for c, txt in enumerate(headers_pres, start=1):
            cell = ws.cell(row=r, column=c, value=txt)
            cell.font = th_font; cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center"); cell.border = border
        r += 1

        if presentaciones_info:
            for (nombre_pres, size), info in sorted(presentaciones_info.items()):
                ws.cell(row=r, column=1, value=nombre_pres)
                ws.cell(row=r, column=2, value=size)
                ws.cell(row=r, column=3, value=info['cajas'])
                ws.cell(row=r, column=4, value=round(info['dinero'], 2))
                for c in range(1, 5):
                    ws.cell(row=r, column=c).border = border
                    ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
                r += 1
        else:
            ws.cell(row=r, column=1, value="(Sin datos)"); r += 1

        r += 1
        ws.cell(row=r, column=1, value="Número total de cajas:").font = h_font
        ws.cell(row=r, column=2, value=total_cajas); r += 1
        ws.cell(row=r, column=1, value="Total equivalente en 11 lbs:").font = h_font
        ws.cell(row=r, column=2, value=round(total_eq_11lbs, 2)); r += 1
        ws.cell(row=r, column=1, value="Total de dinero:").font = h_font
        ws.cell(row=r, column=2, value=round(total_dinero, 2)); r += 2

        ws.cell(row=r, column=1, value="Detalle de embarques de la semana").font = h_font
        r += 1
        headers_det = ["Fecha", "N# Embarque", "N# Factura", "Presentación", "Tamaño", "Cantidad", "Importe"]
        for c, txt in enumerate(headers_det, start=1):
            cell = ws.cell(row=r, column=c, value=txt)
            cell.font = th_font; cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center"); cell.border = border
        r += 1

        for shipment in embarques:
            for item in shipment.items.all():
                importe = item.quantity * float(item.presentation.price)
                ws.cell(row=r, column=1, value=shipment.date.strftime("%d/%m/%Y"))
                ws.cell(row=r, column=2, value=shipment.tracking_number)
                ws.cell(row=r, column=3, value=shipment.invoice_number)
                ws.cell(row=r, column=4, value=item.presentation.name)
                ws.cell(row=r, column=5, value=item.size)
                ws.cell(row=r, column=6, value=item.quantity)
                ws.cell(row=r, column=7, value=round(importe, 2))
                for c in range(1, 7 + 1):
                    ws.cell(row=r, column=c).border = border
                    ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
                r += 1

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 26
        ws.column_dimensions['E'].width = 10
        ws.column_dimensions['F'].width = 20.29
        ws.column_dimensions['G'].width = 14
        ws.column_dimensions['L'].width = 20.29

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"resumen_semana_{year}_W{week:02d}.xlsx"
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response

    # --------------------------
    # Descarga por RANGO (XLSX)
    # --------------------------
    if descargar == 'rango':
        if not request.user.has_perm('empaques.can_download_reports'):
            return HttpResponse("No tienes permiso para descargar reportes.", status=403)

        start_str = request.GET.get('start')
        end_str   = request.GET.get('end')
        if not start_str or not end_str:
            return HttpResponse("Debes indicar 'start' y 'end' en formato YYYY-MM-DD.", status=400)

        try:
            start_d = date.fromisoformat(start_str)
            end_d   = date.fromisoformat(end_str)
        except ValueError:
            return HttpResponse("Fechas inválidas. Usa formato YYYY-MM-DD.", status=400)

        if end_d < start_d:
            return HttpResponse("El fin de rango no puede ser menor que el inicio.", status=400)

        embarques = Shipment.objects.filter(date__range=(start_d, end_d))

        wb = build_summary_xlsx(
            "Resumen de Embarques – Rango de fechas",
            f"Rango: {start_d.strftime('%d/%m/%Y')} – {end_d.strftime('%d/%m/%Y')}",
            embarques
        )

        output = BytesIO(); wb.save(output); output.seek(0)
        filename = f"resumen_rango_{start_d.strftime('%Y%m%d')}_{end_d.strftime('%Y%m%d')}.xlsx"
        resp = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        resp['Content-Disposition'] = f'attachment; filename="{filename}"'
        return resp

    # ================================
    # Descarga mensual o anual (XLSX)
    # ================================
    if descargar in ('mes', 'ano'):
        if descargar == 'mes':
            embarques = (
                Shipment.objects
                .filter(date__year=year, date__month=month)
                .order_by('date', 'tracking_number')
                .prefetch_related('items', 'items__presentation')
            )
            filename = f"resumen_mes_{year}_{month:02d}.xlsx"
            titulo = f"Resumen Mensual {year}-{month:02d}"
        else:
            embarques = (
                Shipment.objects
                .filter(date__year=year)
                .order_by('date', 'tracking_number')
                .prefetch_related('items', 'items__presentation')
            )
            filename = f"resumen_anual_{year}.xlsx"
            titulo = f"Resumen Anual {year}"

        # Recolecta ítems
        items = [it for s in embarques for it in s.items.all()]

        # Agregados
        presentaciones_info = defaultdict(lambda: {'cajas': 0, 'dinero': 0.0})
        total_cajas = 0
        total_eq_11lbs = 0.0
        total_dinero = 0.0

        for it in items:
            key = (it.presentation.name, it.size)
            presentaciones_info[key]['cajas'] += it.quantity
            importe = it.quantity * float(it.presentation.price)
            presentaciones_info[key]['dinero'] += importe
            total_cajas += it.quantity
            total_eq_11lbs += it.quantity * float(it.presentation.conversion_factor)
            total_dinero += importe

        # ==== Construir Excel bonito ====
        wb = Workbook()
        ws = wb.active
        ws.title = "Resumen"

        title_font = Font(name='Calibri', size=18, bold=True, color="3C78D8")
        header_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
        normal_font = Font(name='Calibri', size=12)
        th_fill = PatternFill("solid", fgColor="225577")
        border = Border(
            left=Side(style='thin', color='AAAAAA'),
            right=Side(style='thin', color='AAAAAA'),
            top=Side(style='thin', color='AAAAAA'),
            bottom=Side(style='thin', color='AAAAAA'),
        )

        row = 1
        ws.cell(row=row, column=1, value=titulo).font = title_font
        row += 2

        ws.cell(row=row, column=1, value="Total de Embarques:").font = Font(bold=True)
        ws.cell(row=row, column=2, value=len(embarques)).font = normal_font
        row += 1
        ws.cell(row=row, column=1, value="Número total de cajas:").font = Font(bold=True)
        ws.cell(row=row, column=2, value=total_cajas).font = normal_font
        row += 1
        ws.cell(row=row, column=1, value="Total equivalente en 11 lbs:").font = Font(bold=True)
        ws.cell(row=row, column=2, value=round(total_eq_11lbs, 2)).font = normal_font
        row += 1
        ws.cell(row=row, column=1, value="Total de dinero:").font = Font(bold=True)
        ws.cell(row=row, column=2, value=round(total_dinero, 2)).font = normal_font
        row += 2

        ws.cell(row=row, column=1, value="Presentaciones utilizadas").font = Font(size=14, bold=True)
        row += 1
        headers = ["Presentación", "Tamaño", "Total de cajas", "Total de dinero"]
        for col, h in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = header_font
            cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        row += 1

        for (nombre_pres, size), info in sorted(presentaciones_info.items()):
            ws.cell(row=row, column=1, value=nombre_pres)
            ws.cell(row=row, column=2, value=size)
            ws.cell(row=row, column=3, value=info['cajas'])
            ws.cell(row=row, column=4, value=round(info['dinero'], 2))
            for col in range(1, 4 + 1):
                c = ws.cell(row=row, column=col)
                c.font = normal_font
                c.alignment = Alignment(horizontal="center")
                c.border = border
            row += 1

        ws.column_dimensions['A'].width = 26
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 16

        row += 2

        ws.cell(row=row, column=1, value="Detalle de embarques").font = Font(size=14, bold=True)
        row += 1
        headers_det = ["Fecha", "N# Embarque", "N# Factura", "Presentación", "Tamaño", "Cantidad", "Equiv. 11 lbs", "Importe ($)"]
        for col, h in enumerate(headers_det, start=1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = header_font
            cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        row += 1

        for s in embarques:
            for it in s.items.all():
                eq  = it.quantity * float(it.presentation.conversion_factor)
                amt = it.quantity * float(it.presentation.price)
                vals = [
                    s.date.strftime("%Y-%m-%d"),
                    s.tracking_number,
                    s.invoice_number,
                    it.presentation.name,
                    it.size,
                    it.quantity,
                    round(eq, 2),
                    round(amt, 2),
                ]
                for col, v in enumerate(vals, start=1):
                    cell = ws.cell(row=row, column=col, value=v)
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = border
                row += 1

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response

    # ================================
    # Render normal (sin descargas)
    # ================================
    return render(request, 'empaques/shipment_list.html', {
        'shipments': shipments,
        'today': date.today(),
        'can_download': can_download,
    })




from datetime import date
from django.utils.text import slugify
from io import BytesIO
import os
from django.conf import settings 
from django.http import HttpResponse
@login_required

def daily_report(request, shipment_id=None):
    ship_id  = request.GET.get('shipment_id') or shipment_id
    tracking = request.GET.get('tracking')
    if not (request.user.has_perm('empaques.view_shipment') or request.user.has_perm('empaques.export_reports')):
        return HttpResponseForbidden("No tienes permiso para ver reportes.")  # <-- usa HttpResponseForbidden

    fmt = request.GET.get('format')
    if fmt and not request.user.has_perm('empaques.export_reports'):
        return HttpResponseForbidden("No tienes permiso para exportar reportes.") 
    from io import BytesIO
    import os
    from datetime import date
    from django.conf import settings
    from django.utils.text import slugify
    from django.http import HttpResponse
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    ORDER_FIELD_BY_SLUG = {
        # nombres cortos
        'la-cima-produce':                 'order_lacima',
        'rc-organics':                     'order_rc',
        'gh-farms':                     'order_gh',
        'gourmet-baja-farms':              'order_gourmet',
        'gbf-farms':                       'order_gbf',

        # variantes largas (razón social completa)
        'la-cima-produce-s-p-r-de-r-l':                'order_lacima',
        'rc-organics-s-de-r-l-de-c-v':                 'order_rc',
        'empaque-n-1-s-de-r-l-de-c-v':                 'order_gh',
        'gourmet-baja-farms-s-de-r-l-de-c-v':          'order_gourmet',
        'gbf-farms-s-de-r-l-de-c-v':                   'order_gbf',
    }

    def get_client_order_number(embarque, cliente):
        """Devuelve el número de orden específico del cliente si existe."""
        cname = (cliente or "").lower()
        if "cima" in cname:
            return getattr(embarque, "order_lacima", None)
        if "rc" in cname:
            return getattr(embarque, "order_rc", None)
        if "gourmet" in cname:
            return getattr(embarque, "order_gourmet", None)
        if "gbf" in cname:
            return getattr(embarque, "order_gbf", None)
        if "gh" in cname:
            return getattr(embarque, "order_gh", None)
        return None
    def fecha_es(d):
        """Devuelve la fecha en español: LUNES 31 DE AGOSTO DEL 2025 (en mayúsculas)."""
        if not d:
            return ""
    # Opción A: Babel (recomendado)
        try:
            from babel.dates import format_date
            txt = format_date(d, format="EEEE d 'DE' MMMM 'DEL' y", locale='es_MX')
            return txt.upper()
        except Exception:
            # Opción B: fallback manual si Babel no está disponible
            dias = ["LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO","DOMINGO"]
            meses = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
            return f"{dias[d.weekday()]} {d.day:02d} DE {meses[d.month-1]} DEL {d.year}"
    

    # ---- Lista de clientes ----
    clientes = [ 
        "La Cima Produce",
        "RC Organics",
        "GH Farms",
        "Gourmet Baja Farms",
        "GBF Farms",
    ]
    LEGAL_CLIENT_NAME = {
    "La Cima Produce": "La Cima Produce, S.P.R. DE R.L",
    "RC Organics": "RC Organics S. DE R.L DE C.V.",
    "GH Farms": "Empaque N.1 S. DE R.L. DE C.V.",  
    "Gourmet Baja Farms": "Gourmet Baja Farms S. DE R.L. DE C.V.",
    "GBF Farms": "GBF Farms S. DE R.L. DE C.V.",
}

    clientes_slug = [(c, slugify(c)) for c in clientes] 


    # ---- Fecha a reportar ----
    qdate = request.GET.get('date')
    try:
        report_date = date.fromisoformat(qdate) if qdate else date.today()
    except ValueError:
        report_date = date.today()

    if ship_id:
        try:
            sid = int(ship_id)
            qs = (
                Shipment.objects
                .filter(id=sid)
                .prefetch_related('items', 'items__presentation')
            )
        except ValueError:
            qs = Shipment.objects.none()
    elif tracking:
        qs = (
            Shipment.objects
            .filter(tracking_number=tracking)
            .prefetch_related('items', 'items__presentation')
        )
    else:
        qs = (
            Shipment.objects
            .filter(date=report_date)
            .order_by('-id')
            .prefetch_related('items', 'items__presentation')
        )

    # Si vino selección específica, ajusta report_date para mostrarlo bonito
    if (ship_id or tracking) and qs.exists():
        report_date = qs.first().date
     # Si me pasaron shipment_id o tracking, fuerzo a 1 solo embarque
    if shipment_id or tracking:
        base_qs = Shipment.objects.all().prefetch_related('items', 'items__presentation')
        if shipment_id:
            qs = base_qs.filter(id=shipment_id)
        else:
            qs = base_qs.filter(tracking_number=tracking).order_by('-id')[:1]
        # ajusta la fecha para que los títulos/encabezados muestren el día real del embarque
        if qs:
            report_date = qs[0].date
    

    # Totales del día (para el general) - precomputar total_boxes = sum(item.quantity)
    total_boxes = sum(item.quantity for s in qs for item in s.items.all()) 
    total_eq_11lbs = sum(
        item.quantity * float(item.presentation.conversion_factor)
        for s in qs for item in s.items.all()
    )
    total_amount = sum(
        item.quantity * float(item.presentation.price)
        for s in qs for item in s.items.all()
    )

    # ---------------- Helpers ----------------
    def _str(v):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "Sí" if v else "No" 
        return str(v)

    def write_shipment_info(ws, start_row, start_col, embarque, include_peco=False):
        """Escribe bloque con datos del embarque. Devuelve el último row usado."""
        label_font = Font(name='Calibri', size=18, bold=True, color="666666")
        value_font = Font(name='Calibri', size=16)
        seals = ", ".join([s for s in [embarque.seal_1, embarque.seal_2, embarque.seal_3, embarque.seal_4] if s])
        info = [
            ("Núm. Orden",     _str(embarque.tracking_number)),
            ("FECHA", fecha_es(getattr(embarque, "date", None))),
            ("Transportista",  _str(embarque.carrier)),
            ("Placas Tractor", _str(embarque.tractor_plates)),
            ("Placas Caja",    _str(embarque.box_plates)),
            ("Operador",       _str(embarque.driver)),
            ("Hora Salida",    _str(embarque.departure_time)),
            ("Caja",           _str(embarque.box)),
            ("Cond. de Caja",  _str(embarque.box_conditions)),
            ("Caja sin Olores",_str(embarque.box_free_of_odors)),
            ("Ryan",           _str(embarque.ryan)),
            ("Sellos",         seals),
            ("Chismógrafo",    _str(embarque.chismografo)),
            ("Firma Entrega",  _str(embarque.delivery_signature)),
            ("Firma Operador", _str(embarque.driver_signature)),
            ("Factura",        _str(embarque.invoice_number)),
            ("Dirección",      "H. GALEANA N. 85 LOC A-B, C. ZARAGOZA"),
            ("Y R. ZAPATA COL. CENTRO, 23600", ""),
        ]
         # Insertar Tarimas PECO SOLO si include_peco=True 
        if include_peco:
            info.append(("Tarimas PECO", _str(getattr(embarque, "tarimas_peco", None))))


         # Escribir en hoja
        r = start_row
        for label, value in info:
            ws.cell(row=r, column=start_col,     value=label + ":").font = label_font
            ws.cell(row=r, column=start_col + 1, value=value).font     = value_font
            r += 1

        return r - 1

    def score_shipment(s):
        """Para el general: prioriza más ítems, más campos llenos, más reciente."""
        items_count = len(s.items.all())
        header_values = [
            s.box_conditions, s.box_free_of_odors, s.ryan,
            s.seal_1, s.seal_2, s.seal_3, s.seal_4,
            s.chismografo, s.delivery_signature, s.driver_signature, s.invoice_number
        ]
        filled_count = sum(1 for v in header_values if v not in (None, "", False))
        return (items_count, filled_count, s.id)

    def tarima_temp_text(items_list, tarima):
        """Busca la primera temperatura no vacía en esa tarima y devuelve texto pretty."""
        for it in items_list:
            if it.tarima == tarima and it.temperatura not in (None, ""):
                try:
                    return f"{float(it.temperatura):.1f} °F"
                except Exception:
                    return str(it.temperatura)
        return ""
    

    def pintar_bloque_tarima(ws_, top_row, left_col, temp_col, items_, temp_text):
        
        from openpyxl.styles import Alignment, Border, Side, Font

        thin  = Side(style='thin',   color='999999')
        thick = Side(style='medium', color='000000')
        thick_all = Border(top=thick, bottom=thick, left=thick, right=thick)

        # Marco 2x4 (bordes internos finos, exteriores gruesos)
        for rr in (top_row, top_row + 1):
            for cc in range(left_col, left_col + 4):
                cell = ws_.cell(row=rr, column=cc, value="")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                top_side    = thick if rr == top_row else thin
                bottom_side = thick if rr == (top_row + 1) else thin
                left_side   = thick if cc == left_col else thin
                right_side  = thick if cc == (left_col + 3) else thin
                cell.border = Border(top=top_side, bottom=bottom_side, left=left_side, right=right_side)

        # Celda única de temperatura (fusionada verticalmente)
        ws_.merge_cells(start_row=top_row, start_column=temp_col, end_row=top_row + 1, end_column=temp_col)
        for rr in (top_row, top_row + 1):
            c = ws_.cell(row=rr, column=temp_col, value=(temp_text or "") if rr == top_row else None)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thick_all

        # Hasta 4 ítems (slots: (fila, col_label, col_qty))
        slots = [
            (top_row,     left_col,     left_col + 1),
            (top_row,     left_col + 2, left_col + 3),
            (top_row + 1, left_col,     left_col + 1),
            (top_row + 1, left_col + 2, left_col + 3),
        ]

        # Un poco más alto para que quepan 2 líneas
        ws_.row_dimensions[top_row].height     = 32
        ws_.row_dimensions[top_row + 1].height = 32

        for it, (r, c_label, c_qty) in zip(items_[:4], slots):
            # Texto con salto de línea: Tipo + Tamaño
            label_text = f"{_str(it.presentation.name)}\n{_str(it.size)}"
            lbl = ws_.cell(row=r, column=c_label, value=label_text)
            lbl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            qty = ws_.cell(row=r, column=c_qty, value=it.quantity)
            qty.alignment = Alignment(horizontal="center", vertical="center")
            qty.font = Font(bold=True)

    def write_datos(ws_, start_row, embarque, order_override=None):
        from openpyxl.styles import Font, Alignment
        lf = Font(name='Calibri', size=12, bold=True, color="000000")
        vf = Font(name='Calibri', size=13)

        r = start_row
        if not embarque:
            return r - 1

        # 1) elegir el número de orden a mostrar (cliente > general)
        #    usamos "no vacío" (None o "" cae al general)
        if order_override is not None and str(order_override).strip() != "":
            orden_val = str(order_override).strip()
        else:
            orden_val = _str(embarque.tracking_number)

        # 2) texto de fecha en español (usa tu helper si existe)
        try:
            fecha_txt = fecha_es(getattr(embarque, "date", None))
        except NameError:
            fecha_txt = embarque.date.strftime("%A %d DE %B DEL %Y").upper() if getattr(embarque, "date", None) else ""

        seals = ", ".join([s for s in [embarque.seal_1, embarque.seal_2, embarque.seal_3, embarque.seal_4] if s])

        info = [
            ("NUM. DE ORDEN", orden_val),
            ("FECHA",         fecha_txt),
            ("TRANSPORTISTA", _str(embarque.carrier)),
            ("PLACAS TRACTOR", _str(embarque.tractor_plates)),
            ("PLACAS CAJA",    _str(embarque.box_plates)),
            ("OPERADOR",       _str(embarque.driver)),
            ("HORA DE SALIDA", _str(embarque.departure_time)),
            ("CAJA",           _str(embarque.box)),
            ("CONDICIONES DE LA CAJA", _str(embarque.box_conditions)),
            ("CAJA LIBRE DE OLORES",   _str(embarque.box_free_of_odors)),
            ("RYAN",           _str(embarque.ryan)),
            ("SELLOS",         seals),
            ("CHISMÓGRAFO",    _str(embarque.chismografo)),
            ("FIRMA DEL QUE ENTREGA", _str(embarque.delivery_signature)),
            ("FIRMA DEL OPERADOR",    _str(embarque.driver_signature)),
            ("DEBERÁ MANTENERSE UNA TEMPERATURA CONTINUA DE 35°F", ""),
            ("T. PECO",        _str(getattr(embarque, "tarimas_peco", None))),
            ("DIRECCIÓN",      "H. GALEANA N. 85 LOC A-B, C. ZARAGOZA"),
            ("Y R. ZAPATA COL. CENTRO, 23600", ""),
            ("TELÉFONO",       "01 (613) 132-19-08"),
            ("CEL",            "613 111-71-87, 613 122-01-05"),
        ]

        for label, value in info:
            lbl_up = (label or "").upper()

            # Dirección: valor en B..C (fusionado) y etiqueta en A
            if lbl_up.startswith("DIRECCIÓN"):
                ws_.cell(row=r, column=1, value=f"{label}:").font = lf  # A
                ws_.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
                vcell = ws_.cell(row=r, column=2, value=value)         # B (superior-izq del merge)
                vcell.font = vf
                vcell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                r += 1
                continue

            # 2ª línea de dirección sin etiqueta, solo valor en B..C
            if lbl_up.startswith("Y R. ZAPATA"):
                ws_.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
                vcell = ws_.cell(row=r, column=2, value=value)
                vcell.font = vf
                vcell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                r += 1
                continue

            # Resto normal: etiqueta en A, valor en C
            ws_.cell(row=r, column=1, value=label + ":").font = lf  # A
            ws_.cell(row=r, column=3, value=value).font = vf        # C
            r += 1

        return r - 1


    # =================== EXCEL GENERAL (diseño) ===================
    if request.GET.get('format') == 'xlsx':
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte General"

        header_font = Font(name='Calibri', size=16, bold=True, color="3C78D8")
        table_header_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF")
        th_fill = PatternFill("solid", fgColor="225577")
        border = Border(
            left=Side(style='thin', color='AAAAAA'),
            right=Side(style='thin', color='AAAAAA'),
            top=Side(style='thin', color='AAAAAA'),
            bottom=Side(style='thin', color='AAAAAA'),
        )

        row = 1
        ws.cell(row=row, column=1, value=f"Reporte General – {report_date.strftime('%d/%m/%Y')}")
        ws.cell(row=row, column=1).font = header_font
        row += 2

        # Elegir embarque representativo
        ships = list(qs)
        if ships:
            rep = max(ships, key=score_shipment)
            last = write_shipment_info(ws, start_row=row, start_col=1, embarque=rep)
            row = last + 2
            
        # Encabezados tabla
        headers = [
            "N# EMBARQUE", "N# FACTURA", "PRESENTACIÓN", "TAMAÑO",
            "CANTIDAD", "EQUIV. 11 LBS", "IMPORTE ($)", "CLIENTE"
        ]
        for col, val in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = table_header_font
            cell.fill = th_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        row += 1
         

        # Filas
        for s in qs:
            for item in s.items.all():
                eq = item.quantity * float(item.presentation.conversion_factor)
                amt = item.quantity * float(item.presentation.price)
                ws.cell(row=row, column=1, value=_str(s.tracking_number))
                ws.cell(row=row, column=2, value=_str(s.invoice_number))
                ws.cell(row=row, column=3, value=_str(item.presentation.name))
                ws.cell(row=row, column=4, value=_str(item.size))
                ws.cell(row=row, column=5, value=item.quantity)
                ws.cell(row=row, column=6, value=round(eq, 2))
                ws.cell(row=row, column=7, value=round(amt, 2))
                ws.cell(row=row, column=8, value=_str(item.cliente))
                for c in range(1, 9):
                    ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")
                    ws.cell(row=row, column=c).border = border
                row += 1
                

        # Totales
        ws.merge_cells(start_row=row+1, start_column=1, end_row=row+1, end_column=4)
        ws.cell(row=row+1, column=1, value="TOTALES:").alignment = Alignment(horizontal="right")
        ws.cell(row=row+1, column=1).font = Font(bold=True, color="225577")
        ws.cell(row=row+1, column=5, value=total_boxes)
        ws.cell(row=row+1, column=6, value=round(total_eq_11lbs, 2))
        ws.cell(row=row+1, column=7, value=round(total_amount, 2))
        for c in range(1, 9):
            ws.cell(row=row+1, column=c).font = Font(bold=True)
            ws.cell(row=row+1, column=c).fill = PatternFill("solid", fgColor="BBDDFF")
            ws.cell(row=row+1, column=c).alignment = Alignment(horizontal="center")

        # Anchos
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['H'].width = 20
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"reporte_{report_date}.xlsx"
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
         
    fmt = (request.GET.get('format') or "").strip().lower()
    if fmt.startswith('xlsx_'):
        cliente_slug = fmt[5:]
        cliente = next((c for c in clientes if slugify(c) == cliente_slug), None)
        if not cliente:
            return HttpResponse("Cliente no válido para este reporte.", status=400)

        wb = Workbook()
        ws = wb.active
        ws.title = slugify(cliente)[:31] or "Cliente"

        # --- Anchos A..R (exactos) ---
        widths = {
            'A': 11.43, 'B': 8.57,  'C': 8.57,  'D':10.29, 'E': 7.14,
            'F': 5.43,  'G': 6.14,  'H': 5.57,  'I': 6.86, 'J': 2.86,
            'K': 5.86,  'L': 5.71,  'M': 5.43,  'N': 6.29, 'O': 6.86,
            'P': 3.57,  'Q': 5.00,  'R': 1.71,
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

        # --- Alturas (todas 18.75) ---
        for r in range(1, 80):
            ws.row_dimensions[r].height = 18.75

        # --- Logo en A1 (escalado) ---
        logo_path = os.path.join(settings.BASE_DIR, 'static', 'logos', f'{slugify(cliente)}.png')
        if os.path.exists(logo_path):
            img = XLImage(logo_path)
            target_h = 140
            scale = target_h / img.height
            img.width  = int(img.width * scale)
            img.height = int(img.height * scale)
            ws.add_image(img, "A1")

        # --- Título y subtítulo en I1/I2 ---
        title_font = Font(name="Calibri", size=12, bold=True)
        subtitle_font = Font(name="Calibri", size=11, bold=True)
        display_name = LEGAL_CLIENT_NAME.get(cliente, cliente)

        c1 = ws.cell(row=1, column=9, value=display_name)  # I1
        c1.font = title_font
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2 = ws.cell(row=2, column=9, value="ESPÁRRAGO MANIFIESTO DE EMBARQUE")  # I2
        c2.font = subtitle_font
        c2.alignment = Alignment(horizontal="left", vertical="center")

        # === Selección del embarque del CLIENTE (solo UNO) ===
        shipments_cliente = [s for s in qs if s.items.filter(cliente=cliente).exists()]

        ship_id  = request.GET.get('shipment_id')
        tracking = request.GET.get('tracking')

        rep_cli = None
        # 1) Prioriza ?shipment_id= (si viene numérico)
        if ship_id:
            try:
                sid = int(ship_id)
            except (TypeError, ValueError):
                sid = None
            if sid is not None:
                rep_cli = next((s for s in shipments_cliente if s.id == sid), None)

        # 2) Si no hubo match por id y viene ?tracking=, intenta por tracking_number
        if rep_cli is None and tracking:
            rep_cli = next((s for s in shipments_cliente if _str(s.tracking_number) == tracking), None)

        # 3) Fallback: si no se especificó nada, usa el primero de ese cliente en la fecha
        if rep_cli is None and shipments_cliente:
            rep_cli = shipments_cliente[0]


        # --- Cabecera "EMBARQUE:" (P2:Q2 etiqueta, R2 valor) con NÚMERO POR CLIENTE ---
        label_font = Font(name="Calibri", size=10, bold=True)
        num_font   = Font(name="Calibri", size=10, bold=True, color="FF0000")
        ws.merge_cells(start_row=2, start_column=16, end_row=2, end_column=17)  # P2:Q2
        ws.cell(row=2, column=16, value="EMBARQUE:").font = label_font

        numero_final = ""
        if rep_cli:
            num_cli = get_client_order_number(rep_cli, cliente)  # puede ser None/""
            numero_final = num_cli or _str(rep_cli.tracking_number)
        ws.cell(row=2, column=18, value=numero_final).font = num_font  # R2

        numero_final = ""
        if rep_cli:
            num_cli = get_client_order_number(rep_cli, cliente)  # puede ser None/""
            numero_final = num_cli or _str(rep_cli.tracking_number)
        ws.cell(row=2, column=18, value=numero_final).font = num_font  # R2

        # --- Datos del embarque (izquierda), SOLO UNA VEZ ---
        datos_start_row = 7
        rptr = datos_start_row
        if rep_cli:
            orden_cliente = get_client_order_number(rep_cli, cliente)
            rptr = write_datos(ws, rptr, rep_cli, order_override=orden_cliente) + 2

        # --- GRID de tarimas ---
        grid_start_row = 5
        number_font = Font(name='Calibri', size=8, bold=True, color="444444")
        thick = Side(style='medium', color='000000')
        thick_all = Border(top=thick, bottom=thick, left=thick, right=thick)

        # Ítems SOLO del embarque rep_cli
        items_cliente = list(rep_cli.items.filter(cliente=cliente)) if rep_cli else []

        def temp_txt(tarima_n):
            for it in items_cliente:
                if it.tarima == tarima_n and it.temperatura not in (None, ""):
                    try:
                        return f"{float(it.temperatura):.1f}°F"
                    except Exception:
                        return str(it.temperatura)
            return ""

        # Ajustes de columnas “recorridos 1 a la izquierda”
        ws.column_dimensions['T'].width = 9.57
        ws.column_dimensions['B'].width = 14.86
        ws.column_dimensions['G'].width = 5.86
        ws.column_dimensions['F'].width = 8.86
        ws.column_dimensions['K'].width = 9.57
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['Q'].width = 9.57
        ws.column_dimensions['J'].width = 22
        ws.column_dimensions['H'].width = 22
        ws.column_dimensions['M'].width = 22
        ws.column_dimensions['I'].width = 10.29
        ws.column_dimensions['P'].width = 10.29
        ws.column_dimensions['L'].width = 9.57
        ws.column_dimensions['N'].width = 10.29
        ws.column_dimensions['R'].width = 5.86
        ws.column_dimensions['S'].width = 5.71
        ws.column_dimensions['O'].width = 22
        for rr in range(5, 57):
            ws.row_dimensions[rr].height = 32.25

        for i in range(13):
            top = grid_start_row + i * 2
            t_impar = 1 + 2*i
            t_par   = 2 + 2*i

            num_left_col  = 7   # G
            block_left    = 8   # H..K
            temp_left_l   = 12  # L..M
            block_right   = 13  # N..Q
            temp_right_l  = 17  # R..S
            num_right_col = 18  # T

            # número izquierdo
            for rr in (top, top + 1):
                c = ws.cell(row=rr, column=num_left_col)
                c.font = number_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thick_all
            ws.merge_cells(start_row=top, start_column=num_left_col, end_row=top + 1, end_column=num_left_col)
            ws.cell(row=top, column=num_left_col, value=str(t_impar))

            # bloque izquierdo + temp
            items_impar = [it for it in items_cliente if it.tarima == t_impar][:4]
            pintar_bloque_tarima(ws, top, block_left, temp_left_l, items_impar, temp_txt(t_impar))

            # bloque derecho + temp
            items_par   = [it for it in items_cliente if it.tarima == t_par][:4]
            pintar_bloque_tarima(ws, top, block_right, temp_right_l, items_par, temp_txt(t_par))

            # número derecho
            for rr in (top, top + 1):
                c = ws.cell(row=rr, column=num_right_col)
                c.font = number_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thick_all
            ws.merge_cells(start_row=top, start_column=num_right_col, end_row=top + 1, end_column=num_right_col)
            ws.cell(row=top, column=num_right_col, value=str(t_par))

        # ===== Resumen inferior =====
        from collections import defaultdict
        th_font = Font(name='Calibri', size=16, bold=True, color="FFFFFF")
        th_fill = PatternFill("solid", fgColor="225577")
        thin_border = Border(
            left=Side(style='thin', color='AAAAAA'),
            right=Side(style='thin', color='AAAAAA'),
            top=Side(style='thin', color='AAAAAA'),
            bottom=Side(style='thin', color='AAAAAA'),
        )
        body_font = Font(name='Calibri', size=14)

        grid_last_row = grid_start_row + 13*2 - 1
        data_block_last_row = (rptr - 1) if rep_cli else (datos_start_row - 1)
        summary_top = max(grid_last_row, data_block_last_row) + 2

        # Agregar/agrup ar solo items del rep_cli:
        presentaciones_info = defaultdict(lambda: {'cajas': 0, 'eq11': 0.0})
        total_cajas = 0
        total_eq11  = 0.0

        for it in items_cliente:
            k = (it.presentation.name, it.size)
            presentaciones_info[k]['cajas'] += it.quantity
            eq = it.quantity * float(it.presentation.conversion_factor)
            presentaciones_info[k]['eq11']  += eq
            total_cajas += it.quantity
            total_eq11  += eq

        def merge_pair(row, c1, c2, value=None, *, font=None, fill=None, border=None):
            for cc in range(c1, c2 + 1):
                cell = ws.cell(row=row, column=cc)
                if font:  cell.font = font
                if fill:  cell.fill = fill
                if border: cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
            if value is not None:
                ws.cell(row=row, column=c1, value=value)

        # Encabezados (2 col por campo)
        header_pairs = [(1,2), (3,4), (5,6), (7,8)]
        headers = ["Presentación", "Tamaño", "Cantidad", "Equiv. 11 lbs"]
        for (c1, c2), txt in zip(header_pairs, headers):
            merge_pair(summary_top, c1, c2, txt, font=th_font, fill=th_fill, border=thin_border)

        # Filas
        r = summary_top + 1
        if presentaciones_info:
            for (pres, size), info in sorted(presentaciones_info.items(),
                                     key=lambda kv: (kv[0][0].lower(), str(kv[0][1]).lower())):
                merge_pair(r, 1, 2, pres,               font=body_font, border=thin_border)
                merge_pair(r, 3, 4, size,               font=body_font, border=thin_border)
                merge_pair(r, 5, 6, info['cajas'],      font=body_font, border=thin_border)
                merge_pair(r, 7, 8, round(info['eq11'],2), font=body_font, border=thin_border)
                r += 1
        else:
            merge_pair(r, 1, 8, "(Sin datos)", font=body_font, border=thin_border)
            r += 1

        # Totales
        r += 1
        tot_fill = PatternFill("solid", fgColor="BBDDFF")
        bold = Font(bold=True)
        merge_pair(r, 1, 2, "TOTALES", font=bold, fill=tot_fill, border=thin_border)
        merge_pair(r, 3, 4, "",        font=bold, fill=tot_fill, border=thin_border)
        merge_pair(r, 5, 6, total_cajas,        font=bold, fill=tot_fill, border=thin_border)
        merge_pair(r, 7, 8, round(total_eq11,2), font=bold, fill=tot_fill, border=thin_border)

        # --- Exportar ---
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"reporte_{report_date}_{slugify(cliente)}.xlsx"
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response


    # ---- HTML normal ----- 


    for s in qs:
        for item in s.items.all():
            item.eq_11lbs_calc = item.quantity * float(item.presentation.conversion_factor)
            item.amount_calc   = item.quantity * float(item.presentation.price)

    return render(request, 'empaques/daily_report.html', {
        'report_date': report_date,
        'shipments': qs,
        'total_boxes': total_boxes,
        'total_eq_11lbs': total_eq_11lbs,
        'total_amount': total_amount,
        'clientes': clientes_slug,
    })