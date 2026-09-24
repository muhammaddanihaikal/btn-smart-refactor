import docx

doc_path = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_-_UT_Corporate_Banking_CBD.docx'
doc = docx.Document(doc_path)
for i, tbl in enumerate(doc.tables[4:7]):
    for r in tbl.rows:
        for c in r.cells:
            if 'STATUS' in c.text or 'Success' in c.text or 'Hasil yang diharapkan' in c.text:
                print(f"--- Tbl {i} ---")
                print(c.text)
