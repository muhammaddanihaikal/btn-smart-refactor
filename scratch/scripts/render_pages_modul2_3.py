import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_dir = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816'

doc = fitz.open(pdf_path)

pages_to_render = [
    (13, 'modul_02_page_14.png'),  # 0-indexed 13 is page 14 (TC 2.1 & 2.2)
    (14, 'modul_02_page_15.png'),  # page 15 (TC 2.3)
    (32, 'modul_03_page_33.png'),  # page 33 (TC 3.1)
    (33, 'modul_03_page_34.png')   # page 34 (TC 3.2)
]

for p_idx, filename in pages_to_render:
    page = doc.load_page(p_idx)
    pix = page.get_pixmap(dpi=150)
    out_file = os.path.join(out_dir, filename)
    pix.save(out_file)
    print(f"Saved {filename}")
