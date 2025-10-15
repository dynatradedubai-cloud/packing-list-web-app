import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Packing List Generator", layout="wide")
st.title("📦 Packing List Generator")

uploaded_file = st.file_uploader("Upload Dump Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    # Create a new workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List"

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    def apply_border(ws, cell_range):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = thin_border

    # Header formatting
    ws.merge_cells('A1:J1')
    ws['A1'] = "PACKING LIST"
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].font = Font(bold=True, size=14)

    ws.merge_cells('A2:C2')
    ws.merge_cells('A3:C3')

    invoice_numbers = df['REF1'].dropna().unique()
    ws.merge_cells('E2:J2')
    ws['E2'] = f"Invoice Number: {', '.join(map(str, invoice_numbers))}"
    ws['E2'].alignment = Alignment(horizontal='left', vertical='center')

    ws.merge_cells('E3:J3')
    ws['E3'] = f"Date: {datetime.today().strftime('%d/%m/%Y')}"
    ws['E3'].alignment = Alignment(horizontal='left', vertical='center')

    apply_border(ws, 'A1:J3')

    # Table headers
    headers = ["Sl. No", "CARTONNO", "PARTNO", "QTY", "REF1", "PARTDESC", "WEIGHT", "MANFPART", "CRTN WEIGHT", "Brand"]
    ws.append(headers)
    for col in range(1, 11):
        ws.cell(row=5, column=col).font = Font(bold=True)
        ws.cell(row=5, column=col).alignment = Alignment(horizontal='center', vertical='center')
    apply_border(ws, 'A5:J5')

    # Grouping and merging logic
    grouped = df.groupby(['CARTONNO', 'CRTN WEIGHT'])
    start_row = 6
    sl_no = 1
    for (carton, weight), group in grouped:
        rows = group.to_dict('records')
        for i, row in enumerate(rows):
            ws.append([
                sl_no if i == 0 else "",
                carton if i == 0 else "",
                row['PARTNO'],
                row['QTY'],
                row['REF1'],
                row['PARTDESC'],
                row['WEIGHT'],
                row['MANFPART'],
                weight if i == 0 else "",
                row['Brand']
            ])
        end_row = start_row + len(rows) - 1
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        ws.merge_cells(start_row=start_row, start_column=9, end_row=end_row, end_column=9)
        for r in range(start_row, end_row + 1):
            for c in range(1, 11):
                ws.cell(row=r, column=c).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=r, column=c).border = thin_border
        start_row = end_row + 1
        sl_no += 1

    # Footer formatting
    total_qty = df['QTY'].sum()
    total_weight = df['CRTN WEIGHT'].drop_duplicates().sum()
    package_count = df['CARTONNO'].nunique()

    ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=3)
    ws.cell(row=start_row, column=2).value = "Total Quantity"
    ws.cell(row=start_row, column=2).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=start_row, column=4).value = total_qty
    ws.cell(row=start_row, column=4).alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells(start_row=start_row, start_column=6, end_row=start_row, end_column=8)
    ws.cell(row=start_row, column=6).value = "TOTAL CARTON WEIGHT"
    ws.cell(row=start_row, column=6).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(row=start_row, column=9).value = total_weight
    ws.cell(row=start_row, column=9).alignment = Alignment(horizontal='center', vertical='center')

    apply_border(ws, f'B{start_row}:C{start_row}')
    apply_border(ws, f'F{start_row}:H{start_row}')

    start_row += 2
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=10)
    ws.cell(row=start_row, column=1).value = f"NO OF PACKAGES : {package_count}"
    ws.cell(row=start_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
    apply_border(ws, f'A{start_row}:J{start_row}')

    start_row += 1
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=10)
    ws.cell(row=start_row, column=1).value = f"TOTAL GROSS WEIGHT : {round(total_weight)} KG"
    ws.cell(row=start_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
    apply_border(ws, f'A{start_row}:J{start_row}')

    start_row += 2
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+3, end_column=10)
    ws.cell(row=start_row, column=1).value = "AUTHORISED SIGNATORY"
    ws.cell(row=start_row, column=1).alignment = Alignment(horizontal='left', vertical='bottom')
    apply_border(ws, f'A{start_row}:J{start_row+3}')

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)

    st.success("Packing List created successfully!")
    st.download_button(
        label="📥 Download Packing List",
        data=output.getvalue(),
        file_name="Packing_List.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
