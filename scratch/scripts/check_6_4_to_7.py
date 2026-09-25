import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

tcs = ['6.4', '6.5', '6.6', '6.7']

for i, tbl in enumerate(doc.tables):
    r0 = tbl.rows[0].cells[0].text.strip()
    if r0 in tcs:
        print(f"Table {i}: TC {r0}")
        print(f"  Expected: {repr(tbl.rows[2].cells[1].text.strip())}")
