import io
import pandas as pd
from typing import Dict, Any

def generate_excel_dashboard(dfs: Dict[str, pd.DataFrame], report: Dict[str, Any], date_param: str) -> bytes:
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # 1. CREATE DASHBOARD SHEET
        dash = workbook.add_worksheet("DASHBOARD GERENCIA")
        dash.hide_gridlines(2)
        dash.set_column('A:A', 3)
        dash.set_column('B:D', 28)
        dash.set_column('E:H', 22)
        
        # Formats
        title_fmt = workbook.add_format({'bold': True, 'font_size': 22, 'font_color': '#ffffff', 'bg_color': '#1e293b', 'align': 'center', 'valign': 'vcenter'})
        header_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#334155', 'font_color': '#ffffff', 'bottom': 2, 'bottom_color': '#ef9b00'})
        kpi_label_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'font_color': '#64748b', 'align': 'left'})
        kpi_val_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'num_format': '#,##0.00 "EUR"', 'font_color': '#0f172a', 'align': 'left'})
        kpi_bad_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'num_format': '#,##0.00 "EUR"', 'font_color': '#ef4444', 'align': 'left'})
        kpi_good_fmt = workbook.add_format({'bold': True, 'font_size': 20, 'num_format': '#,##0.00 "EUR"', 'font_color': '#10b981', 'align': 'left'})
        
        alert_title_fmt = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#b91c1c', 'bg_color': '#fef2f2', 'border': 1, 'border_color': '#f87171'})
        alert_text_fmt = workbook.add_format({'font_size': 11, 'font_color': '#7f1d1d', 'bg_color': '#fef2f2', 'border': 1, 'border_color': '#f87171', 'text_wrap': True})
        
        table_label_fmt = workbook.add_format({'font_size': 12, 'font_color': '#334155'})
        table_val_fmt = workbook.add_format({'font_size': 12, 'num_format': '#,##0.00 "EUR"', 'font_color': '#0f172a'})
        table_total_fmt = workbook.add_format({'bold': True, 'font_size': 13, 'num_format': '#,##0.00 "EUR"', 'font_color': '#0f172a', 'top': 1})
        
        # Extract KPIs
        month_invoiced = report.get("month_invoiced", 0.0)
        forecast_total = report.get("forecast_total", 0.0)
        budget = report.get("forecast", {}).get("budget", report.get("budget", 2150256.00))
        gap = report.get("budget_gap_expected", forecast_total - budget)
        
        # Title
        dash.merge_range('B2:H3', f'RESUMEN EJECUTIVO COMERCIAL - {date_param}', title_fmt)
        
        # High Level KPIs
        dash.write('B5', 'FACTURADO MES ACTUAL', kpi_label_fmt)
        dash.write('B6', month_invoiced, kpi_val_fmt)
        
        dash.write('C5', 'PREVISIÓN DE CIERRE', kpi_label_fmt)
        dash.write('C6', forecast_total, kpi_val_fmt)
        
        dash.write('D5', 'DESVIACIÓN vs PRESUPUESTO', kpi_label_fmt)
        dash.write('D6', gap, kpi_good_fmt if gap >= 0 else kpi_bad_fmt)
        
        # Breakdown Table
        dash.write('B9', 'COMPOSICIÓN DEL CIERRE ESTIMADO', header_fmt)
        dash.write('B10', '1. Facturado Real', table_label_fmt)
        dash.write('C10', month_invoiced, table_val_fmt)
        
        pending_albaranes = report.get("pending_delivery_amount", 0.0)
        dash.write('B11', '2. Albaranes Pendientes', table_label_fmt)
        dash.write('C11', pending_albaranes, table_val_fmt)
        
        loadable = report.get("loadable_amount", 0.0)
        dash.write('B12', '3. Pedidos Cargables (>7 días)', table_label_fmt)
        dash.write('C12', loadable, table_val_fmt)
        
        older = report.get("older_amount", 0.0)
        dash.write('B13', '4. Pedidos Antiguos (>1 mes)', table_label_fmt)
        dash.write('C13', older, table_val_fmt)
        
        dash.write('B14', 'TOTAL ESTIMACIÓN DE CIERRE', workbook.add_format({'bold': True, 'font_size': 13}))
        dash.write('C14', forecast_total, table_total_fmt)
        
        # Add Chart
        chart = workbook.add_chart({'type': 'doughnut'})
        
        # We need to write the chart data to a hidden sheet or somewhere safe
        data_sheet = workbook.add_worksheet('ChartData')
        data_sheet.hide()
        data_sheet.write_column('A1', ['Facturado', 'Albaranes', 'Pedidos Recientes', 'Pedidos Antiguos'])
        data_sheet.write_column('B1', [month_invoiced, pending_albaranes, loadable, older])
        
        chart.add_series({
            'name': 'Composición',
            'categories': ['ChartData', 0, 0, 3, 0],
            'values':     ['ChartData', 0, 1, 3, 1],
            'points': [
                {'fill': {'color': '#1e293b'}},
                {'fill': {'color': '#ef9b00'}},
                {'fill': {'color': '#3b82f6'}},
                {'fill': {'color': '#ef4444'}}
            ],
            'data_labels': {'percentage': True, 'font': {'color': '#ffffff', 'bold': True}}
        })
        chart.set_title({'name': 'Distribución del Cierre', 'name_font': {'size': 12, 'bold': True}})
        chart.set_legend({'position': 'bottom'})
        chart.set_style(10)
        chart.set_size({'width': 480, 'height': 300})
        dash.insert_chart('E8', chart)
        
        # Alerts & Recommendations
        dash.write('B17', 'ALERTAS Y RECOMENDACIONES CLAVE PARA GERENCIA', header_fmt)
        row = 18
        alerts = report.get("recommendations", [])
        if not alerts:
            dash.merge_range(f'B{row}:H{row}', "✅ No hay alertas críticas para el día de hoy.", table_label_fmt)
        else:
            for rec in alerts:
                dash.merge_range(f'B{row}:H{row}', f"⚠️ {rec}", alert_text_fmt)
                row += 1
                dash.set_row(row-1, 30) # make row taller
                
        # Status Orders
        dash.write('B24', 'ESTADO DE PEDIDOS (ZONA)', header_fmt)
        dash.write('B25', 'Zona', workbook.add_format({'bold': True, 'bottom': 1}))
        dash.write('C25', 'Estado', workbook.add_format({'bold': True, 'bottom': 1}))
        dash.write('D25', 'Importe', workbook.add_format({'bold': True, 'bottom': 1, 'align': 'right'}))
        
        row = 26
        for st in report.get("status_summary", []):
            dash.write(f'B{row}', st["zona"])
            dash.write(f'C{row}', st["estado"])
            dash.write(f'D{row}', st["importe"], table_val_fmt)
            row += 1

        # 2. DATA SHEETS
        for name, df in dfs.items():
            if df.empty:
                continue
            sheet_name = name.capitalize()
            # Remove timezones for Excel
            df_export = df.copy()
            for col in df_export.columns:
                if pd.api.types.is_datetime64_any_dtype(df_export[col]):
                    df_export[col] = df_export[col].dt.tz_localize(None)
                    
            df_export.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
            worksheet = writer.sheets[sheet_name]
            
            (max_row, max_col) = df_export.shape
            column_settings = [{'header': c} for c in df_export.columns]
            
            # Find a nice table style
            style = 'Table Style Medium 2' # Blue
            if name == 'facturas': style = 'Table Style Medium 4' # Green
            elif name == 'ofertas': style = 'Table Style Medium 3' # Orange
            
            worksheet.add_table(0, 0, max_row, max_col - 1, {
                'columns': column_settings,
                'style': style,
                'autofilter': True
            })
            # Adjust column widths
            worksheet.set_column(0, max_col - 1, 16)
            
            # Format numeric columns
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            for i, col in enumerate(df_export.columns):
                if 'importe' in col.lower() or 'precio' in col.lower() or 'total' in col.lower():
                    worksheet.set_column(i, i, 16, money_fmt)
                    
    return output.getvalue()
