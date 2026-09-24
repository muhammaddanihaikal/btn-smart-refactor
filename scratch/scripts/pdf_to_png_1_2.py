import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_path = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816\tc_1_2_preview.png'

doc = fitz.open(pdf_path)
target_page = 0
for i in range(len(doc)):
    text = doc[i].get_text()
    if '1.2' in text and 'Melakukan login dengan email dan password valid' in text:
        target_page = i
        print(f"Found TC 1.2 on page {i+1}")
        break

page = doc.load_page(target_page)
pix = page.get_pixmap(dpi=150)
pix.save(out_path)
print(f'Saved preview to {out_path}')
