import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_dir = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816'

doc = fitz.open(pdf_path)

pages_found = []
for i in range(len(doc)):
    text = doc[i].get_text()
    if '5.1' in text and 'Clock In' in text:
        pages_found.append(i)
    elif '5.2' in text and 'Clock In' in text and i not in pages_found:
        pages_found.append(i)

print(f"Pages found for Modul 05: {[p+1 for p in pages_found]}")
for idx, p_num in enumerate(pages_found):
    page = doc.load_page(p_num)
    pix = page.get_pixmap(dpi=150)
    out_file = os.path.join(out_dir, f'modul_05_page_{p_num+1}.png')
    pix.save(out_file)
    print(f"Saved {out_file}")
