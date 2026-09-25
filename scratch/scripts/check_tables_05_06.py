import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

tcs = ['5.1', '5.2', '5.3', '5.4', '5.5', '5.6', '5.7', '6.1', '6.2', '6.3']

for i, tbl in enumerate(doc.tables):
    r0 = tbl.rows[0].cells[0].text.strip()
    tc_no = r0
    is_first = False
    if r0 == 'No.' and len(tbl.rows) > 1:
        tc_no = tbl.rows[1].cells[0].text.strip()
        is_first = True
        
    if tc_no in tcs:
        print(f"Table {i}: TC {tc_no} (is_first={is_first})")
