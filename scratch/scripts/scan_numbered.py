import docx
import re

doc_path = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

found_numbered = []

for i, tbl in enumerate(doc.tables):
    for r_idx, r in enumerate(tbl.rows):
        for c_idx, c in enumerate(r.cells):
            text = c.text
            if "Hasil yang diharapkan" in text:
                # check if there is numbering like "1. " or "2. "
                if re.search(r'(?:^|\n)\s*1\.\s+', text):
                    found_numbered.append((i, r_idx, c_idx, tbl.rows[0].cells[0].text.strip(), text.strip()))

print(f"Total tables with numbered expected results: {len(found_numbered)}")
for item in found_numbered[:20]:
    print(f"Table {item[0]} (TC {item[3]}):")
    print(f"  {repr(item[4])}\n")
