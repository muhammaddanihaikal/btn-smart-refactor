import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_dir = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816'

doc = fitz.open(pdf_path)

targets = [
    ('Membuka halaman profile', 'tc_2_1_preview.png'),
    ('Informasi Pribadi', 'tc_2_2_preview.png'),
    ('Mengubah foto profil', 'tc_2_3_preview.png'),
    ('3.1', 'tc_3_1_preview.png')
]

for kw, filename in targets:
    for i in range(len(doc)):
        text = doc[i].get_text()
        if kw in text:
            print(f"Keyword '{kw}' found on page {i+1}")
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=150)
            out_file = os.path.join(out_dir, filename)
            pix.save(out_file)
            print(f"Saved {filename}")
            break
