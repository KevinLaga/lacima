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


@login_required
@permission_required('empaques.add_shipment', raise_exception=True)

def shipment_create(request):
    """
    Vista para capturar un nuevo embarque con sus ítems.
    """
    ItemFormSet = inlineformset_factory(
        Shipment,
        ShipmentItem,
        form=ShipmentItemForm,
        formset=BaseShipmentItemFormSet,
        extra=26,
        can_delete=True,
    )

    if request.method == 'POST':
        form = ShipmentForm(request.POST)
        formset = ItemFormSet(request.POST)
        if form.is_valid() and formset.is_valid():
            shipment = form.save()
            formset.instance = shipment
            formset.save()
            return redirect('shipment_list')
    else:
        form = ShipmentForm()
        formset = ItemFormSet()

    return render(request, 'empaques/shipment_form.html', {
        'form': form,
        'formset': formset,
    })
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from datetime import date
from collections import defaultdict

from .models import Shipment, ShipmentItem

@login_required
def shipment_list(request):
    # Imports locales para la exportación
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

    # Solo aceptamos descargas si tiene permiso
    descargar = request.GET.get('descargar') if can_download else None

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
        items = []
        for s in embarques:
            for it in s.items.all():
                items.append(it)

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

        # Estilos
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

        # Totales globales (bloque)
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

        # Tabla: Presentaciones utilizadas
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
            for col in range(1, 5):
                c = ws.cell(row=row, column=col)
                c.font = normal_font
                c.alignment = Alignment(horizontal="center")
                c.border = border
            row += 1

        # Ancho columnas
        ws.column_dimensions['A'].width = 26
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 16

        row += 2

        # Detalle por embarque
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

        # Descargar
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

def daily_report(request):
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

    # ---- Lista de clientes ----
    clientes = [ 
        "La Cima Produce",
        "RC Organics",
        "GH Farms",
        "Gourmet Baja Farms",
        "GBF Farms",
    ]
    clientes_slug = [(c, slugify(c)) for c in clientes] 

    # ---- Fecha a reportar ----
    qdate = request.GET.get('date')
    try:
        report_date = date.fromisoformat(qdate) if qdate else date.today() 
    except ValueError:
        report_date = date.today()

    # Prefetch y orden -id (más reciente primero)
    qs = (
        Shipment.objects
        .filter(date=report_date)
        .order_by('-id')
        .prefetch_related('items', 'items__presentation')
    )
    

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

    def write_shipment_info(ws, start_row, start_col, embarque): 
        """Escribe bloque con datos del embarque. Devuelve el último row usado."""
        label_font = Font(name='Calibri', size=13, bold=True, color="666666")
        value_font = Font(name='Calibri', size=13)
        seals = ", ".join([s for s in [embarque.seal_1, embarque.seal_2, embarque.seal_3, embarque.seal_4] if s])
        info = [
            ("Núm. Orden",     _str(embarque.tracking_number)),
            ("Fecha",          embarque.date.strftime("%d-%m-%Y") if getattr(embarque, "date", None) else ""),
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
        ]
        r = start_row 

        for label, val in info:
            ws.cell(row=r, column=start_col,     value=label + ":").font = label_font
            ws.cell(row=r, column=start_col + 1, value=val).font        = value_font
            ws.cell(row=r, column=start_col).alignment     = Alignment(horizontal="left")
            ws.cell(row=r, column=start_col + 1).alignment = Alignment(horizontal="left") 
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
        """Dibuja bloque 2x4 (tipo/tam/cant | tipo/tam/cant) + celda única de temperatura."""
        thin  = Side(style='thin',   color='999999')
        thick = Side(style='medium', color='000000')
        thick_all = Border(top=thick, bottom=thick, left=thick, right=thick)

        # 8 celdas internas (2 filas x 4 columnas) con marco grueso exterior
        for rr in (top_row, top_row + 1):
            for cc in range(left_col, left_col + 4):
                cell = ws_.cell(row=rr, column=cc, value="")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                top_side    = thick if rr == top_row else thin
                bottom_side = thick if rr == (top_row + 1) else thin
                left_side   = thick if cc == left_col else thin
                right_side  = thick if cc == (left_col + 3) else thin
                cell.border = Border(top=top_side, bottom=bottom_side, left=left_side, right=right_side)

        # Celda única de temperatura (fusionada verticalmente) → aplicar borde a ambas filas
        ws_.merge_cells(start_row=top_row, start_column=temp_col, end_row=top_row + 1, end_column=temp_col)
        for rr in (top_row, top_row + 1):
            c = ws_.cell(row=rr, column=temp_col, value=(temp_text or "") if rr == top_row else None)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thick_all

        # Pinta hasta 2 ítems (izquierda y derecha del bloque)
        if not items_:
            return
        it1 = items_[0]
        ws_.cell(row=top_row,     column=left_col,     value=_str(it1.presentation.name))
        ws_.cell(row=top_row + 1, column=left_col,     value=_str(it1.size))
        ws_.cell(row=top_row + 1, column=left_col + 1, value=it1.quantity)

        if len(items_) >= 2:
            it2 = items_[1]
            ws_.cell(row=top_row,     column=left_col + 2, value=_str(it2.presentation.name))
            ws_.cell(row=top_row + 1, column=left_col + 2, value=_str(it2.size))
            ws_.cell(row=top_row + 1, column=left_col + 3, value=it2.quantity)

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

    # =================== EXCEL POR CLIENTE (grid con temperatura) ===================
    for cliente in clientes:
        if request.GET.get('format') == f'xlsx_{slugify(cliente)}':
            wb = Workbook()
            ws = wb.active

            # Logo y posiciones base
            logo_path = os.path.join(settings.BASE_DIR, 'static', 'logos', f'{slugify(cliente)}.png')
            if os.path.exists(logo_path):
                img = XLImage(logo_path)
                img.height = 120
                img.width  = 260
                ws.add_image(img, "A1")
                grid_start_row  = 6   # grid arriba (no choca con logo)
                datos_start_row = 12  # datos debajo del logo
            else:
                grid_start_row  = 4
                datos_start_row = 8

            # Embarques que sí tienen items de este cliente
            shipments_cliente = [s for s in qs if s.items.filter(cliente=cliente).exists()]

            # Datos del embarque (a la izquierda)
            rptr = datos_start_row
            for embarque in shipments_cliente:
                last = write_shipment_info(ws, start_row=rptr, start_col=2, embarque=embarque)
                rptr = last + 2
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 50

            # ---- GRID UNIFICADO (tarimas) a la derecha de los datos ----
            number_font = Font(name='Calibri', size=12, bold=True, color="444444")
            # Bordes gruesos para celdas fusionadas (número y temperatura)
            thick = Side(style='medium', color='000000')
            thick_all = Border(top=thick, bottom=thick, left=thick, right=thick)

            base_col = 5  # columna E para pegarlo a los datos

            # Ítems de este cliente (para temperaturas por tarima)
            items_cliente = [item for s in qs for item in s.items.filter(cliente=cliente)]

            def set_col_width(col_idx, width):
                ws.column_dimensions[get_column_letter(col_idx)].width = width

            number_col_width = 4.0
            data_col_width   = 5.43
            temp_col_width   = 6.24

            for i in range(13):
                top = grid_start_row + i * 2
                tarima_impar = 1 + 2*i
                tarima_par   = 2 + 2*i

                # [num impar] [bloque impar x4] [temp impar] [bloque par x4] [temp par] [num par]
                num_left_col    = base_col
                left_block_col  = num_left_col + 1 
                left_temp_col   = left_block_col + 4 
                right_block_col = left_temp_col + 1
                right_temp_col  = right_block_col + 4
                num_right_col   = right_temp_col + 1

                # Anchos de columnas
                set_col_width(num_left_col,  number_col_width)
                for cc in range(left_block_col, left_block_col + 4):
                    set_col_width(cc, data_col_width)
                set_col_width(left_temp_col,  temp_col_width)
                for cc in range(right_block_col, right_block_col + 4):
                    set_col_width(cc, data_col_width)
                set_col_width(right_temp_col, temp_col_width)
                set_col_width(num_right_col,  number_col_width)

                # Ítems por tarima (máximo 2 por bloque)
                items_impar = [it for it in items_cliente if it.tarima == tarima_impar][:2]
                items_par   = [it for it in items_cliente if it.tarima == tarima_par][:2]

                # Temperaturas por tarima (texto listo)
                temp_left_text  = tarima_temp_text(items_cliente, tarima_impar)
                temp_right_text = tarima_temp_text(items_cliente, tarima_par)

                # NÚMERO IZQUIERDA (tarima impar) – celda fusionada 2 filas, borde completo
                ws.merge_cells(start_row=top, end_row=top+1, start_column=num_left_col, end_column=num_left_col)
                for rr in (top, top+1):
                    c = ws.cell(row=rr, column=num_left_col, value=str(tarima_impar) if rr == top else None)
                    c.font = number_font
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = thick_all

                # Bloques con temperatura (función ya corrige borde de la celda de temperatura)
                pintar_bloque_tarima(ws, top, left_block_col,  left_temp_col,  items_impar, temp_left_text)
                pintar_bloque_tarima(ws, top, right_block_col, right_temp_col, items_par,   temp_right_text)

                # NÚMERO DERECHA (tarima par) – celda fusionada 2 filas, borde completo
                ws.merge_cells(start_row=top, end_row=top+1, start_column=num_right_col, end_column=num_right_col)
                for rr in (top, top+1):
                    c = ws.cell(row=rr, column=num_right_col, value=str(tarima_par) if rr == top else None)
                    c.font = number_font
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.border = thick_all

            # ---------- Tabla resumen inferior (SIN temperatura) ---------- 
            table_header_font = Font(name='Calibri', size=14, bold=True, color="FFFFFF") 
            th_fill = PatternFill("solid", fgColor="225577")
            border_thin = Border(
                left=Side(style='thin', color='AAAAAA'),
                right=Side(style='thin', color='AAAAAA'),
                top=Side(style='thin', color='AAAAAA'),
                bottom=Side(style='thin', color='AAAAAA'),
            )

            # Usa el último rptr como fin del bloque de datos
            data_block_last_row = rptr - 1 if shipments_cliente else (datos_start_row - 1)
            grid_last_row = grid_start_row + 13*2 - 1
            after_grid_row = max(data_block_last_row, grid_last_row) + 2

            headers = [
                "N# EMBARQUE", "N# FACTURA", "PRESENTACIÓN", "TAMAÑO",
                "CANTIDAD", "EQUIV. 11 LBS", "IMPORTE ($)"
            ]
            for idx, texto in enumerate(headers, 1):
                cell = ws.cell(row=after_grid_row, column=idx, value=texto)
                cell.font = table_header_font
                cell.fill = th_fill
                cell.alignment = Alignment(horizontal="center")
                cell.border = border_thin

            row = after_grid_row + 1
            cliente_total_boxes = 0
            cliente_total_eq_11lbs = 0
            cliente_total_amount = 0
            for s in qs:
                for item in s.items.filter(cliente=cliente):
                    eq = item.quantity * float(item.presentation.conversion_factor)
                    amt = item.quantity * float(item.presentation.price)
                    ws.cell(row=row, column=1, value=_str(s.tracking_number))
                    ws.cell(row=row, column=2, value=_str(s.invoice_number))
                    ws.cell(row=row, column=3, value=_str(item.presentation.name))
                    ws.cell(row=row, column=4, value=_str(item.size))
                    ws.cell(row=row, column=5, value=item.quantity)
                    ws.cell(row=row, column=6, value=round(eq, 2))
                    ws.cell(row=row, column=7, value=round(amt, 2))
                    for c in range(1, 8):
                        ws.cell(row=row, column=c).border = border_thin
                        ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")
                    cliente_total_boxes += item.quantity
                    cliente_total_eq_11lbs += eq
                    cliente_total_amount += amt
                    row += 1

            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            ws.cell(row=row, column=1, value="TOTALES:").alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=1).font = Font(bold=True, color="225577")
            ws.cell(row=row, column=5, value=cliente_total_boxes)
            ws.cell(row=row, column=6, value=round(cliente_total_eq_11lbs, 2))
            ws.cell(row=row, column=7, value=round(cliente_total_amount, 2))
            for c in range(1, 8):
                ws.cell(row=row, column=c).font = Font(bold=True)
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor="BBDDFF")
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")

            # Descargar
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
