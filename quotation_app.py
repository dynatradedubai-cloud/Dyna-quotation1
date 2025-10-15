import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

def format_excel(df):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws_main = wb.active
    ws_main.title = "Quotation"
    ws_zero = wb.create_sheet("Zero Stock")
    ws_yellow = wb.create_sheet("Yellow Card")

    fill = PatternFill(start_color="FFF6D6", end_color="FFF6D6", fill_type="solid")
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    def write_headers(ws):
        ws.merge_cells("A1:I1")
        ws["A1"] = "DYNATRADE AUTOMOTIVE LLC - QUOTATION"
        ws.merge_cells("A2:B2")
        ws["A2"] = "Customer Code"
        ws.merge_cells("C2:E2")
        ws["C2"] = ""
        ws.merge_cells("A3:B3")
        ws["A3"] = "Customer Name"
        ws.merge_cells("C3:E3")
        ws["C3"] = ""
        ws.merge_cells("F2:G3")
        ws["F2"] = "Date:"
        ws.merge_cells("H2:I3")
        ws["H2"] = datetime.today().strftime("%d/%m/%Y")

        for row in [1, 2, 3]:
            for col in range(1, 10):
                cell = ws.cell(row=row, column=col)
                cell.fill = fill
                cell.font = bold_font
                cell.alignment = center_align
                cell.border = thin_border

        headers = ["S.No", "Inquired Part No", "Part Number", "Manf.Part", "Description", "Brand", "Stock on Hand", "Unit Price", "COO"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_num)
            cell.value = header
            cell.fill = fill
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = thin_border

    def write_data(ws, data):
        current_row = 5
        previous_serial = None
        for _, row in data.iterrows():
            serial = row["S.No"]
            if previous_serial is not None and serial != previous_serial:
                current_row += 1
            for col_num, value in enumerate(row, 1):
                cell = ws.cell(row=current_row, column=col_num, value=value)
                cell.border = thin_border
            previous_serial = serial
            current_row += 1

    write_headers(ws_main)
    write_headers(ws_zero)
    write_headers(ws_yellow)

    quotation_data = df[(df["Stock on Hand"] != 0) & (df["Brand"].notna())]
    zero_stock_data = df[df["Stock on Hand"] == 0]
    yellow_card_data = df[df["Brand"].isna()]
    common_serials = set(quotation_data["S.No"])
    zero_stock_data = zero_stock_data[~zero_stock_data["S.No"].isin(common_serials)]

    write_data(ws_main, quotation_data)
    write_data(ws_zero, zero_stock_data)
    write_data(ws_yellow, yellow_card_data)

    for ws in [ws_main, ws_zero, ws_yellow]:
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

    wb.save(output)
    return output

st.title("Dynatrade Quotation Formatter")
uploaded_file = st.file_uploader("Upload Dump Excel File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    columns = ["S.No", "Inquired Part No", "Part Number", "Manf.Part", "Description", "Brand", "Stock on Hand", "Unit Price", "COO"]
    df_filtered = df[columns].copy()
    formatted_file = format_excel(df_filtered)
    st.download_button("Download Formatted Quotation", formatted_file.getvalue(), file_name="Formatted_Quotation.xlsx")
