import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

# Define header style
header_fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
header_font = Font(bold=True)
header_alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))

# Function to apply header formatting
def format_headers(ws):
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

    for cell in ["A1", "A2", "C2", "A3", "C3", "F2", "H2"]:
        ws[cell].fill = header_fill
        ws[cell].font = header_font
        ws[cell].alignment = header_alignment
        ws[cell].border = thin_border

    headers = ["S.No", "Inquired Part No", "Part Number", "Manf.Part", "Description",
               "Brand", "Stock on Hand", "Unit Price", "COO"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

# Function to autofit column widths
def autofit_columns(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

# Function to write data to sheet
def write_data(ws, df):
    current_row = 5
    last_serial = None
    for _, row in df.iterrows():
        serial = row["S.No"]
        if last_serial is not None and serial != last_serial:
            current_row += 1  # empty row
        for col_num, col_name in enumerate(["S.No", "Inquired Part No", "Part Number", "Manf.Part",
                                            "Description", "Brand", "Stock on Hand", "Unit Price", "COO"], 1):
            cell = ws.cell(row=current_row, column=col_num, value=row[col_name])
            cell.border = thin_border
        last_serial = serial
        current_row += 1

# Function to process and generate formatted workbook
def generate_workbook(df):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Quotation"
    format_headers(ws1)
    valid_df = df[(df["Stock on Hand"] != 0) & (df["Brand"].notna())]
    write_data(ws1, valid_df)
    autofit_columns(ws1)

    ws2 = wb.create_sheet("Zero Stock")
    format_headers(ws2)
    zero_df = df[df["Stock on Hand"] == 0]
    zero_df = zero_df[~zero_df["S.No"].isin(valid_df["S.No"])]
    write_data(ws2, zero_df)
    autofit_columns(ws2)

    ws3 = wb.create_sheet("Yellow Card")
    format_headers(ws3)
    yellow_df = df[df["Brand"].isna()]
    write_data(ws3, yellow_df)
    autofit_columns(ws3)

    return wb

# Streamlit UI
st.title("Quotation Formatter")
uploaded_file = st.file_uploader("Upload Excel Dump", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    wb = generate_workbook(df)
    output = io.BytesIO()
    wb.save(output)
    st.download_button("Download Formatted Quotation", output.getvalue(), file_name="Formatted_Quotation.xlsx")
