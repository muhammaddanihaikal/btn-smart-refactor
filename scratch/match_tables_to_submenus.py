import docx
import openpyxl

wb = openpyxl.load_workbook(r'd:\Project\BTN\SIT\Test Case.xlsx', data_only=True)
ws = wb['TC BTN SMART Web']

# Get list of all rows in Excel: (mod, sub, title)
excel_rows = []
for r in range(2, ws.max_row + 1):
    mod = ws.cell(r, 2).value
    sub = ws.cell(r, 3).value
    title = ws.cell(r, 4).value
    if mod and title:
        excel_rows.append((str(mod).strip(), str(sub).strip() if sub else '', str(title).strip()))

doc = docx.Document(r'H:\My Drive\Zegen\BTN Smart\Refactor\SIT\SIT BTN SMART Web (Updated).docx')

data_tables = [t for t in doc.tables if len(t.columns) == 7 and len(t.rows) > 1]
print(f"Total data tables in Updated doc: {len(data_tables)}")

# For each table, find which module and sub menu its TCs belong to
for i, t in enumerate(data_tables):
    first_title = t.rows[1].cells[1].text.strip()
    last_title = t.rows[-1].cells[1].text.strip()
    
    # Match with excel_rows by title substring or similarity
    matched_sub = []
    matched_mod = []
    for m, s, title in excel_rows:
        if first_title.lower() in title.lower() or title.lower() in first_title.lower():
            matched_mod.append(m)
            matched_sub.append(s)
            break
            
    mod_name = matched_mod[0] if matched_mod else "Unknown"
    sub_name = matched_sub[0] if matched_sub else "Unknown"
    print(f"Tbl {i+1:2d} ({len(t.rows)-1:2d} TCs) -> [{mod_name}] - [{sub_name}] | First: '{first_title[:35]}'")
