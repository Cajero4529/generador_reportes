import os
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import io

app = Flask(__name__)
CORS(app)

class GeneradorReportes:
    def __init__(self):
        self.naranja = 'FF8C00'
        self.bordó = '800020'
        self.gris_claro = 'E8E8E8'
        self.gris_oscuro = 'D3D3D3'
    
    def procesar(self, archivo_bytes, pto_venta, z_codigo):
        """Procesa el archivo y genera el reporte"""
        try:
            wb = load_workbook(io.BytesIO(archivo_bytes))
            ws = wb['Operaciones']
            
            headers = [cell.value for cell in ws[1]]
            datos = []
            
            for row in ws.iter_rows(min_row=2, values_only=False):
                fila = {}
                for idx, cell in enumerate(row):
                    if idx < len(headers):
                        fila[headers[idx]] = cell.value
                datos.append(fila)
            
            tickets_por_caja = {}
            for fila in datos:
                caja = fila.get('N° Caja')
                if caja:
                    if caja not in tickets_por_caja:
                        tickets_por_caja[caja] = []
                    tickets_por_caja[caja].append(fila)
            
            tickets_1art = []
            tickets_2art = []
            
            for caja, articulos in tickets_por_caja.items():
                total = sum(float(art.get('Subtotal', 0) or 0) for art in articulos)
                cupón = str(articulos[0].get('N° Cupón', ''))
                
                if len(articulos) == 1:
                    tickets_1art.append({
                        'caja': caja,
                        'cupón': cupón,
                        'total': total,
                        'articulos': articulos
                    })
                elif len(articulos) == 2:
                    tickets_2art.append({
                        'caja': caja,
                        'cupón': cupón,
                        'total': total,
                        'articulos': articulos
                    })
            
            tickets_1art.sort(key=lambda x: x['total'], reverse=True)
            tickets_2art.sort(key=lambda x: x['total'], reverse=True)
            
            lotes = self._crear_lotes(tickets_1art, tickets_2art)
            excel_bytes = self._generar_excel(lotes, pto_venta, z_codigo)
            
            return excel_bytes
            
        except Exception as e:
            print(f"Error: {str(e)}")
            raise
    
    def _crear_lotes(self, tickets_1art, tickets_2art):
        """Crea lotes respetando distribución Benford"""
        lotes = []
        lote_actual = {'numero': 1, 'cajas': [], 'items': 0, 'monto': 0, 'tickets': []}
        
        idx_1art = 0
        idx_2art = 0
        
        while idx_1art < len(tickets_1art) or idx_2art < len(tickets_2art):
            items_faltantes = 9 - lote_actual['items']
            
            if items_faltantes == 0:
                lotes.append(lote_actual)
                lote_actual = {'numero': len(lotes) + 1, 'cajas': [], 'items': 0, 'monto': 0, 'tickets': []}
                items_faltantes = 9
            
            added = False
            
            if idx_1art < len(tickets_1art) and items_faltantes >= 1:
                ticket = tickets_1art[idx_1art]
                lote_actual['cajas'].append(ticket['caja'])
                lote_actual['items'] += 1
                lote_actual['monto'] += ticket['total']
                lote_actual['tickets'].append(ticket)
                idx_1art += 1
                added = True
            elif idx_2art < len(tickets_2art) and items_faltantes >= 2:
                ticket = tickets_2art[idx_2art]
                lote_actual['cajas'].append(ticket['caja'])
                lote_actual['items'] += 2
                lote_actual['monto'] += ticket['total']
                lote_actual['tickets'].append(ticket)
                idx_2art += 1
                added = True
            
            if not added:
                if lote_actual['cajas']:
                    lotes.append(lote_actual)
                    lote_actual = {'numero': len(lotes) + 1, 'cajas': [], 'items': 0, 'monto': 0, 'tickets': []}
                else:
                    break
        
        if lote_actual['cajas']:
            lotes.append(lote_actual)
        
        return lotes
    
    def _generar_excel(self, lotes, pto_venta, z_codigo):
        """Genera el archivo Excel"""
        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        header_fill = PatternFill(start_color=self.naranja, end_color=self.naranja, fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        lote_header_fill = PatternFill(start_color=self.gris_oscuro, end_color=self.gris_oscuro, fill_type='solid')
        lote_header_font = Font(bold=True, size=11)
        subheader_fill = PatternFill(start_color=self.gris_claro, end_color=self.gris_claro, fill_type='solid')
        subheader_font = Font(bold=True, size=10)
        total_fill = PatternFill(start_color=self.bordó, end_color=self.bordó, fill_type='solid')
        total_font = Font(bold=True, color='FFFFFF', size=11)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        sheet_resumen = wb.create_sheet("Resumen Lotes", 0)
        row = 1
        
        for lote in lotes:
            cajas_sorted = sorted(lote['cajas'], reverse=True)
            
            sheet_resumen.merge_cells(f'A{row}:D{row}')
            cell = sheet_resumen.cell(row=row, column=1)
            cell.value = f"LOTE {lote['numero']} - Cajas: {', '.join(str(c) for c in cajas_sorted)}"
            cell.font = lote_header_font
            cell.fill = lote_header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center')
            sheet_resumen.row_dimensions[row].height = 18
            row += 1
            
            headers = ['Nro.', 'N° Cupón', 'Detalle del Concepto', 'Abona']
            for col, header in enumerate(headers, 1):
                cell = sheet_resumen.cell(row=row, column=col)
                cell.value = header
                cell.font = subheader_font
                cell.fill = subheader_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            row += 1
            
            nro_orden = 1
            for ticket in lote['tickets']:
                for articulo in ticket['articulos']:
                    cell = sheet_resumen.cell(row=row, column=1)
                    cell.value = nro_orden
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                    
                    cell = sheet_resumen.cell(row=row, column=2)
                    cell.value = str(ticket['cupón'])
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                    cell.number_format = '@'
                    
                    codigo = articulo.get('Código', '')
                    cantidad = int(articulo.get('Cantidad', 0) or 0)
                    detalle = f"{pto_venta}/{z_codigo} / {codigo} / {cantidad}"
                    
                    cell = sheet_resumen.cell(row=row, column=3)
                    cell.value = detalle
                    cell.border = border
                    cell.alignment = Alignment(horizontal='left', wrap_text=True)
                    
                    monto = float(articulo.get('Subtotal', 0) or 0)
                    cell = sheet_resumen.cell(row=row, column=4)
                    cell.value = int(monto)
                    cell.number_format = '$#,##0'
                    cell.border = border
                    cell.alignment = Alignment(horizontal='right')
                    
                    nro_orden += 1
                    row += 1
            
            sheet_resumen.merge_cells(f'A{row}:C{row}')
            cell = sheet_resumen.cell(row=row, column=1)
            cell.value = f"TOTAL LOTE {lote['numero']}"
            cell.font = Font(bold=True)
            cell.fill = lote_header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='right')
            
            cell = sheet_resumen.cell(row=row, column=4)
            cell.value = int(lote['monto'])
            cell.number_format = '$#,##0'
            cell.font = Font(bold=True)
            cell.fill = lote_header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='right')
            
            row += 2
        
        sheet_resumen.merge_cells(f'A{row}:C{row}')
        cell = sheet_resumen.cell(row=row, column=1)
        cell.value = "TOTAL GENERAL"
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right')
        
        cell = sheet_resumen.cell(row=row, column=4)
        cell.value = int(sum(l['monto'] for l in lotes))
        cell.number_format = '$#,##0'
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right')
        
        sheet_resumen.column_dimensions['A'].width = 8
        sheet_resumen.column_dimensions['B'].width = 15
        sheet_resumen.column_dimensions['C'].width = 50
        sheet_resumen.column_dimensions['D'].width = 15
        
        sheet_completo = wb.create_sheet("Completo", 1)
        
        sheet_completo['A1'] = 'DETALLE DE PRODUCTOS POR LOTE - DISTRIBUCIÓN NATURAL DE BENFORD'
        sheet_completo['A1'].font = Font(bold=True, size=14, color='FFFFFF')
        sheet_completo['A1'].fill = header_fill
        sheet_completo.merge_cells('A1:E1')
        sheet_completo['A1'].alignment = Alignment(horizontal='center', vertical='center')
        sheet_completo.row_dimensions[1].height = 25
        
        fecha = datetime.now().strftime('%d/%m/%Y %H:%M')
        sheet_completo['A2'] = f'Procesado: {fecha} | Pto. Venta: {pto_venta} | Z: {z_codigo}'
        sheet_completo['A2'].font = Font(italic=True, size=9)
        sheet_completo.merge_cells('A2:E2')
        
        row = 4
        
        for lote in lotes:
            cajas_sorted = sorted(lote['cajas'], reverse=True)
            
            sheet_completo.merge_cells(f'A{row}:E{row}')
            cell = sheet_completo.cell(row=row, column=1)
            cell.value = f"LOTE {lote['numero']} - Cajas: {', '.join(str(c) for c in cajas_sorted)}"
            cell.font = lote_header_font
            cell.fill = lote_header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center')
            sheet_completo.row_dimensions[row].height = 18
            row += 1
            
            headers_resumen = ['Código', 'Descripción', 'Cantidad', 'Precio Unit.', 'Monto']
            for col, header in enumerate(headers_resumen, 1):
                cell = sheet_completo.cell(row=row, column=col)
                cell.value = header
                cell.font = Font(bold=True, size=10)
                cell.fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            row += 1
            
            for ticket in lote['tickets']:
                for articulo in ticket['articulos']:
                    cell = sheet_completo.cell(row=row, column=1)
                    cell.value = articulo.get('Código', '')
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                    
                    cell = sheet_completo.cell(row=row, column=2)
                    cell.value = articulo.get('Título', '')
                    cell.border = border
                    cell.alignment = Alignment(horizontal='left', wrap_text=True)
                    
                    cell = sheet_completo.cell(row=row, column=3)
                    cell.value = int(articulo.get('Cantidad', 0) or 0)
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center')
                    
                    cell = sheet_completo.cell(row=row, column=4)
                    cell.value = float(articulo.get('Precio Unit.', 0) or 0)
                    cell.number_format = '$#,##0'
                    cell.border = border
                    cell.alignment = Alignment(horizontal='right')
                    
                    cell = sheet_completo.cell(row=row, column=5)
                    cell.value = int(float(articulo.get('Subtotal', 0) or 0))
                    cell.number_format = '$#,##0'
                    cell.border = border
                    cell.alignment = Alignment(horizontal='right')
                    
                    row += 1
            
            total_cantidad = sum(int(art.get('Cantidad', 0) or 0) for ticket in lote['tickets'] for art in ticket['articulos'])
            total_monto = sum(float(art.get('Subtotal', 0) or 0) for ticket in lote['tickets'] for art in ticket['articulos'])
            
            cell = sheet_completo.cell(row=row, column=1)
            cell.value = "SUMA TOTAL"
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
            
            cell = sheet_completo.cell(row=row, column=2)
            cell.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
            cell.border = border
            
            cell = sheet_completo.cell(row=row, column=3)
            cell.value = total_cantidad
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
            
            cell = sheet_completo.cell(row=row, column=4)
            cell.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
            cell.border = border
            
            cell = sheet_completo.cell(row=row, column=5)
            cell.value = int(total_monto)
            cell.number_format = '$#,##0'
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
            cell.border = border
            cell.alignment = Alignment(horizontal='right')
            
            row += 3
        
        sheet_completo.column_dimensions['A'].width = 14
        sheet_completo.column_dimensions['B'].width = 45
        sheet_completo.column_dimensions['C'].width = 12
        sheet_completo.column_dimensions['D'].width = 15
        sheet_completo.column_dimensions['E'].width = 15
        
        excel_file = io.BytesIO()
        wb.save(excel_file)
        excel_file.seek(0)
        return excel_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generar', methods=['POST'])
def generar():
    try:
        if 'archivo' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        archivo = request.files['archivo']
        pto_venta = request.form.get('pto_venta')
        z_codigo = request.form.get('z_codigo')
        
        if not pto_venta or not z_codigo:
            return jsonify({'error': 'Faltan datos'}), 400
        
        archivo_bytes = archivo.read()
        generador = GeneradorReportes()
        excel_bytes = generador.procesar(archivo_bytes, pto_venta, z_codigo)
        
        return send_file(
            excel_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Reporte_Lotes_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
