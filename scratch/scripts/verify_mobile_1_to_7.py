import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

for tbl in doc.tables[:12]:
    r0 = tbl.rows[0].cells[0].text.strip()
    tc_no = r0
    if r0 == 'No.' and len(tbl.rows) > 1:
        tc_no = tbl.rows[1].cells[0].text.strip()
        
    for r in tbl.rows:
        for c in r.cells:
            if "Hasil yang diharapkan" in c.text:
                print(f"[{tc_no}]: {c.text.strip().replace(chr(10), ' | ')}")
                break
