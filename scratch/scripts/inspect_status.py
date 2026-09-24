import docx

doc_path = r'D:\Project\BTN\Hasil Uji\Template Dokumen Hasil Uji.docx'
doc = docx.Document(doc_path)
for i, tbl in enumerate(doc.tables[:6]):
    for r in tbl.rows:
        for c in r.cells:
            if '[STATUS]' in c.text or '[Success]' in c.text or 'Status' in c.text:
                print(f"Tbl {i}: {c.text}")
