import docx

doc = docx.Document(r'D:\Project\BTN Smart\Refactor\SIT\SIT BTN SMART Mobile.docx')

rows_data = []
for t in doc.tables:
    for r in t.rows:
        cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
        if len(cells) >= 4 and any(c.isdigit() for c in cells[0]):
            rows_data.append(cells)

print(f"Total rows in Mobile SIT: {len(rows_data)}")

for i, row in enumerate(rows_data):
    no = row[0]
    title = row[1]
    steps = row[2]
    expected = row[3]
    
    # check comparison with previous row
    if i > 0:
        prev_no = rows_data[i-1][0]
        prev_title = rows_data[i-1][1]
        prev_steps = rows_data[i-1][2]
        
        # 1. Exact step duplicate
        if steps.lower() == prev_steps.lower():
            print(f"[EXACT STEP DUPLICATE]")
            print(f"  Current : {no} - {title}")
            print(f"            Steps: {steps}")
            print(f"  Previous: {prev_no} - {prev_title}")
            print(f"            Steps: {prev_steps}")
            print()
            
        # 2. 'Membuka' followed by 'Menampilkan / Melihat'
        elif 'membuka' in prev_title.lower() and any(k in title.lower() for k in ['menampilkan', 'melihat']) and any(k in title.lower() for k in ['daftar', 'data', 'list']):
            print(f"[REDUNDANT 'MENAMPILKAN DAFTAR' AFTER 'MEMBUKA']")
            print(f"  Current : {no} - {title}")
            print(f"            Steps: {steps}")
            print(f"            Expected: {expected}")
            print(f"  Previous: {prev_no} - {prev_title}")
            print(f"            Steps: {prev_steps}")
            print()
