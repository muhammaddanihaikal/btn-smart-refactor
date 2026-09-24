import docx

doc_path = r'D:\Project\BTN Smart\Refactor\SIT\SIT BTN SMART Mobile.docx'
doc = docx.Document(doc_path)

for tbl in doc.tables:
    for r in tbl.rows:
        c0 = r.cells[0].text.strip()
        if c0 in ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7']:
            print(f"TC {c0}:")
            print(f"  Title: {r.cells[1].text.strip()}")
            if len(r.cells) > 3:
                print(f"  Expected (col 3): {repr(r.cells[3].text.strip())}")
            print()
