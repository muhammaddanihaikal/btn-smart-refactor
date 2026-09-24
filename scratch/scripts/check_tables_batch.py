import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

tcs = [
    '1.12', '1.13', '1.14', '1.15',
    '2.3', '2.4', '2.5', '2.6', '2.7', '2.8', '2.9', '2.10', '2.11', '2.12', '2.13'
]

found = {}
for i, tbl in enumerate(doc.tables):
    r0 = tbl.rows[0].cells[0].text.strip()
    if r0 in tcs:
        found[r0] = i
    elif r0 == 'No.' and len(tbl.rows) > 1:
        r1 = tbl.rows[1].cells[0].text.strip()
        if r1 in tcs:
            found[r1] = i

for tc in tcs:
    if tc in found:
        print(f"TC {tc}: found at Table {found[tc]}")
    else:
        print(f"TC {tc}: NOT FOUND!")
