# Merchant_Statement_Formatter

Files:
- app.py       GUI
- converter.py Cleaning/parsing logic
- requirements.txt Dependencies

The parser is content-based. It does not assume a fixed number of rows, pages, transactions, or header positions.

Supported input:
- .xls
- .xlsx
- .xlsm
- .zip containing Excel files

Output modes:
- Separate
- Merge all
- Group by N files

Run:
1. py -m pip install -r requirements.txt
2. py app.py
