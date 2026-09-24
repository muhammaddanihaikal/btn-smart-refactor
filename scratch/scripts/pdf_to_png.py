import fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_path = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816\tc_1_1_preview.png'

doc = fitz.open(pdf_path)
# Find the page with TC 1.1
target_page = 0
for i in range(len(doc)):
    if '1.1 Membuka halaman Login' in doc[i].get_text():
        target_page = i
        break

page = doc.load_page(target_page)
pix = page.get_pixmap(dpi=150)
pix.save(out_path)
print(f'Saved preview to {out_path}')
