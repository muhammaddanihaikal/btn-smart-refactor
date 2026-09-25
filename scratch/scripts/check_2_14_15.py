import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

for i, tbl in enumerate(doc.tables):
    r0 = tbl.rows[0].cells[0].text.strip()
    if r0 in ['2.14', '2.15']:
        print(f"Table {i}: TC {r0}")
        print(f"  Expected: {repr(tbl.rows[2].cells[1].text.strip())}")
