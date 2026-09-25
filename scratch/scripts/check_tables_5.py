import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

for i, tbl in enumerate(doc.tables):
    r0 = tbl.rows[0].cells[0].text.strip()
    tc_no = r0
    is_first = False
    if r0 == 'No.' and len(tbl.rows) > 1:
        tc_no = tbl.rows[1].cells[0].text.strip()
        is_first = True
        
    if tc_no in ['5.1', '5.2']:
        print(f"Table {i}: TC {tc_no} (is_first={is_first}, rows={len(tbl.rows)}, cols={len(tbl.columns)})")
        if is_first:
            print(f"   Image row: 2, Expected row: 3")
            print(f"   Expected text: {repr(tbl.rows[3].cells[1].text.strip())}")
        else:
            print(f"   Image row: 1, Expected row: 2")
            print(f"   Expected text: {repr(tbl.rows[2].cells[1].text.strip())}")
