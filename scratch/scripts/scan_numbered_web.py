import docx
import re
import os

for path in [
    r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Web.docx',
    r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Web.docx'
]:
    if not os.path.exists(path):
        continue
    print(f"Scanning: {path}")
    doc = docx.Document(path)
    found_numbered = []
    for i, tbl in enumerate(doc.tables):
        for r_idx, r in enumerate(tbl.rows):
            for c_idx, c in enumerate(r.cells):
                text = c.text
                if "Hasil yang diharapkan" in text:
                    if re.search(r'(?:^|\n)\s*1\.\s+', text):
                        found_numbered.append((i, tbl.rows[0].cells[0].text.strip(), text.strip()))
    print(f"  Total tables with numbered expected results: {len(found_numbered)}")
    for item in found_numbered[:10]:
        print(f"  Table {item[0]} (TC {item[1]}): {repr(item[2][:80])}...")
