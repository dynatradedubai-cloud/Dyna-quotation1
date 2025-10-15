Quotation Formatter App

This Streamlit app allows you to upload an Excel dump file and download a formatted quotation workbook with three sheets:

1. Sheet1 - Main quotation with valid stock and brand
2. Zero Stock - Items with zero stock, excluding serials from Sheet1
3. Yellow Card - Items with missing brand

Formatting includes:
- Merged headers with date
- Column headers in row 4
- Light yellow fill, bold font, center alignment
- Thin borders and spacing between serial groups

To run:
pip install streamlit pandas openpyxl
streamlit run app.py
