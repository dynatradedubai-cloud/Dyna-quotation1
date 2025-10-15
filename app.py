import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Styling constants
HEADER_FILL = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
HEADER_FONT = Font(bold=True)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))

# Header setup
HEADER_ROWS = [
    ["DYNATRADE AUTOMOTIVE LLC - QUOTATION"] + [""] * 8,
    ["Customer Code", ""] + [""] * 7,
    ["", "", "", "", ""],
    ["Customer Name", ""] + [""] * 7,
    ["", "", "", "", ""],
    ["Date:", ""] + [""] * 6 + [datetime.datetime.today().strftime("%d/%m/%Y"), ""],
    [""] * 9,
    ["S.No", "Inquired Part No", "Part Number", "Manf.Part", "Description", "Brand",
     "Stock on Hand", "Unit Price", "COO"]
]

MERGE_CELLS = [
    ("A1", "I1"), ("A2", "B2"), ("C2", "E2"),
    ("A3", "B3"), ("C3", "E3"), ("F2", "G3"), ("H2", "I3")
]

def apply_header(ws):
    for i, row in enumerate(HEADER_ROWS, start=1):
        for j, val in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = HEADER_ALIGN
            cell.border = THIN_BORDER
    for start, end in MERGE_CELLS:
        ws.merge_cells(f"{start}:{end}")

def write_data(ws, df):
    current_row = len(HEADER_ROWS) + 1
    prev_serial = None
    for _, row in df.iterrows():
        serial = row['S.No']
        if prev_serial is not None and serial != prev_serial:
            for col in range(1, 10):
                cell = ws.cell(row=current_row, column=col, value="")
                cell.border = THIN_BORDER
            current_row += 1
        for col_index, col_name in enumerate(["S.No", "Inquired Part No", "Part Number", "Manf.Part",
                                              "Description", "Brand", "Stock on Hand", "Unit Price", "COO"], start=1):
            cell = ws.cell(row=current_row, column=col_index, value=row[col_name])
            cell.border = THIN_BORDER
        prev_serial = serial
        current_row += 1

def adjust_column_widths(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.coordinate in ws.merged_cells:
                continue
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

# Streamlit UI
st.title("Quotation Formatter")
uploaded_file = st.file_uploader("Upload Excel Dump File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    df = df[["S.No", "Inquired Part No", "Part Number", "Manf.Part", "Description",
             "Brand", "Stock on Hand", "Unit Price", "COO"]]

    # Create workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    apply_header(ws1)
    df_main = df[(df["Stock on Hand"] != 0) & (df["Brand"].notna())]
    write_data(ws1, df_main)
    adjust_column_widths(ws1)

    # Sheet2: Zero Stock
    df_zero = df[df["Stock on Hand"] == 0]
    df_zero_filtered = df_zero[~df_zero["S.No"].isin(df_main["S.No"])]
    ws2 = wb.create_sheet("Zero Stock")
    apply_header(ws2)
    write_data(ws2, df_zero_filtered)
    adjust_column_widths(ws2)

    # Sheet3: Yellow Card
    df_yellow = df[df["Brand"].isna()]
    ws3 = wb.create_sheet("Yellow Card")
    apply_header(ws3)
    write_data(ws3, df_yellow)
    adjust_column_widths(ws3)

    # Save to buffer
    output = BytesIO()
    wb.save(output)
    st.download_button("Download Formatted Quotation", output.getvalue(), file_name="Formatted_Quotation.xlsx")
