import openpyxl
wb = openpyxl.load_workbook(r'D:\Project\BTN Smart\Refactor\Test Script\Uji Sistem.xlsx', read_only=True)
ws = wb['TC BTN SMART Mobile']
for r in range(1, 5):
    row_vals = []
    for c in range(1, 10):
        row_vals.append(str(ws.cell(r, c).value))
    print(f"Row {r}: {row_vals}")
