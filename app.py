
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font

st.title("Packing List Generator")
uploaded_file = st.file_uploader("Upload Dump Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    required_columns = ['CARTONNO', 'PARTNO', 'QTY', 'REF1', 'PARTDESC', 'WEIGHT', 'MANFPART', 'CRTN WEIGHT', 'Brand']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing columns in uploaded file: {', '.join(missing_columns)}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Packing List"

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

        ws.merge_cells('A1:J1')
        ws['A1'] = "PACKING LIST"
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A1'].font = Font(bold=True)

        ws.merge_cells('A2:C2')
        ws.merge_cells('A3:C3')

        ws.merge_cells('E2:J2')
        invoice_numbers = df['REF1'].dropna().unique()
        ws['E2'] = f"Invoice Number: {', '.join(map(str, invoice_numbers))}"
        ws['E2'].alignment = Alignment(horizontal='left', vertical='center')

        ws.merge_cells('E3:J3')
        ws['E3'] = f"Date: {datetime.now().strftime('%d/%m/%Y')}"
        ws['E3'].alignment = Alignment(horizontal='left', vertical='center')

        headers = ['Sl. No', 'CARTONNO', 'PARTNO', 'QTY', 'REF1', 'PARTDESC', 'WEIGHT', 'MANFPART', 'CRTN WEIGHT', 'Brand']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.border = thin_border

        grouped = df.groupby(['CARTONNO', 'CRTN WEIGHT'])

        start_row = 6
        sl_no = 1
        for (carton, weight), group in grouped:
            num_rows = len(group)
            for i, (_, row) in enumerate(group.iterrows()):
                current_row = start_row + i
                ws.cell(row=current_row, column=3, value=row['PARTNO'])
                ws.cell(row=current_row, column=4, value=row['QTY'])
                ws.cell(row=current_row, column=5, value=row['REF1'])
                ws.cell(row=current_row, column=6, value=row['PARTDESC'])
                ws.cell(row=current_row, column=7, value=row['WEIGHT'])
                ws.cell(row=current_row, column=8, value=row['MANFPART'])
                ws.cell(row=current_row, column=10, value=row['Brand'])
                for col in range(1, 11):
                    ws.cell(row=current_row, column=col).border = thin_border

            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row+num_rows-1, end_column=1)
            ws.cell(row=start_row, column=1, value=sl_no)
            ws.cell(row=start_row, column=1).alignment = Alignment(horizontal='center', vertical='center')

            ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row+num_rows-1, end_column=2)
            ws.cell(row=start_row, column=2, value=carton)
            ws.cell(row=start_row, column=2).alignment = Alignment(horizontal='center', vertical='center')

            ws.merge_cells(start_row=start_row, start_column=9, end_row=start_row+num_rows-1, end_column=9)
            ws.cell(row=start_row, column=9, value=weight)
            ws.cell(row=start_row, column=9).alignment = Alignment(horizontal='center', vertical='center')

            sl_no += 1
            start_row += num_rows

        footer_row = start_row + 1
        ws.merge_cells(start_row=footer_row, start_column=2, end_row=footer_row, end_column=3)
        ws.cell(row=footer_row, column=2, value="Total Quantity")
        ws.cell(row=footer_row, column=2).alignment = Alignment(horizontal='right', vertical='center')
        total_qty = df['QTY'].sum()
        ws.cell(row=footer_row, column=4, value=total_qty)

        ws.merge_cells(start_row=footer_row, start_column=6, end_row=footer_row, end_column=8)
        ws.cell(row=footer_row, column=6, value="TOTAL CARTON WEIGHT")
        ws.cell(row=footer_row, column=6).alignment = Alignment(horizontal='right', vertical='center')
        total_weight = df.drop_duplicates(subset=['CARTONNO'])['CRTN WEIGHT'].sum()
        ws.cell(row=footer_row, column=9, value=total_weight)

        ws.merge_cells(start_row=footer_row+1, start_column=1, end_row=footer_row+1, end_column=10)
        ws.cell(row=footer_row+1, column=1, value=f"NO OF PACKAGES : {df['CARTONNO'].nunique()}")

        ws.merge_cells(start_row=footer_row+2, start_column=1, end_row=footer_row+2, end_column=10)
        ws.cell(row=footer_row+2, column=1, value=f"TOTAL GROSS WEIGHT : {round(total_weight)} KG")

        ws.merge_cells(start_row=footer_row+3, start_column=1, end_row=footer_row+6, end_column=10)
        ws.cell(row=footer_row+3, column=1, value="AUTHORISED SIGNATORY")
        ws.cell(row=footer_row+3, column=1).alignment = Alignment(horizontal='left', vertical='bottom')

        for r in range(footer_row, footer_row+6):
            for c in range(1, 11):
                ws.cell(row=r, column=c).border = thin_border

        output = BytesIO()
        wb.save(output)

        st.success("Packing List created successfully!")
        st.download_button(
            label="Download Packing List",
            data=output.getvalue(),
            file_name="Packing_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
