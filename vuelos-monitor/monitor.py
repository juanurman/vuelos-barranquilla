import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

EXCEL_FILE = "vuelos_rio_ano_nuevo_2027.xlsx"


def save_to_excel(prices):
  now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

  # Si el archivo no existe, lo creamos con encabezados y diseño
  if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Historial Monitoreo"

    # Título principal
    ws.merge_cells("A1:I1")
    ws["A1"] = "Monitoreo Vuelos Río de Janeiro - Año Nuevo 2026/2027"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(
        start_color="1A73E8", end_color="1A73E8", fill_type="solid"
    )
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "Fecha y Hora",
        "30 Dic - 02 Ene (USD)",
        "30 Dic - 03 Ene (USD)",
        "31 Dic - 02 Ene (USD)",
        "31 Dic - 03 Ene (USD)",
        "Solo Ida 30 Dic",
        "Solo Ida 31 Dic",
        "Solo Vuelta 02 Ene",
        "Solo Vuelta 03 Ene",
    ]
    ws.append([])  # Fila 2 vacía
    ws.append(headers)  # Fila 3 encabezados

    # Estilo de encabezados
    header_fill = PatternFill(
        start_color="F1F3F4", end_color="F1F3F4", fill_type="solid"
    )
    for col in range(1, len(headers) + 1):
      cell = ws.cell(row=3, column=col)
      cell.font = Font(name="Arial", size=10, bold=True)
      cell.fill = header_fill
      cell.alignment = Alignment(horizontal="center", vertical="center")
  else:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Historial Monitoreo"]

  # Fila con los datos del chequeo actual
  row_data = [
      now_str,
      prices.get("RIO_30_02", "Sin dato"),
      prices.get("RIO_30_03", "Sin dato"),
      prices.get("RIO_31_02", "Sin dato"),
      prices.get("RIO_31_03", "Sin dato"),
      prices.get("IDA_30", "Sin dato"),
      prices.get("IDA_31", "Sin dato"),
      prices.get("VTA_02", "Sin dato"),
      prices.get("VTA_03", "Sin dato"),
  ]
  ws.append(row_data)

  # Resaltar en verde si algún precio es < 500 USD
  last_row = ws.max_row
  green_fill = PatternFill(
      start_color="D4EDDA", end_color="D4EDDA", fill_type="solid"
  )
  for col in range(2, len(row_data) + 1):
    val = ws.cell(row=last_row, column=col).value
    if isinstance(val, (int, float)) and val < 500:
      ws.cell(row=last_row, column=col).fill = green_fill
      ws.cell(row=last_row, column=col).font = Font(bold=True, color="155724")

  # Ajuste automático del ancho de columnas
  for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

  wb.save(EXCEL_FILE)
  print(f"✅ Excel actualizado localmente en '{EXCEL_FILE}'.")
