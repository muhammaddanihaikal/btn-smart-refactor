import docx

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

for i, tbl in enumerate(doc.tables[:10]):
    print(f"--- Table {i} ---")
    visited = set()
    for r_idx, r in enumerate(tbl.rows):
        for c_idx, c in enumerate(r.cells):
            if c not in visited:
                visited.add(c)
                text = c.text.replace('\n', ' | ')
                print(f"R{r_idx}C{c_idx}: {text}")
