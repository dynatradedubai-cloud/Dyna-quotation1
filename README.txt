Quotation Formatter Streamlit App

Features:
- Upload an Excel dump file.
- Automatically formats the file into three sheets:
  - Sheet1: Main quotation (valid Brand and Stock on Hand > 0)
  - Zero Stock: Items with Stock on Hand = 0 (excluding common serials with Sheet1)
  - Yellow Card: Items with empty Brand

How to Run Locally:
1. Install dependencies:
   pip install streamlit pandas openpyxl

2. Run the app:
   streamlit run app.py

3. Upload your Excel file and download the formatted quotation.
