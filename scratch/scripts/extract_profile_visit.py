import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_dir = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816'

doc = fitz.open(pdf_path)

targets = [
    ('2.1', 'tc_2_1_preview.png'),
    ('2.2', 'tc_2_2_preview.png'),
    ('3.1', 'tc_3_1_preview.png')
]

for tc_id, filename in targets:
    found_page = -1
    for i in range(len(doc)):
        text = doc[i].get_text()
        if f' {tc_id} ' in f' {text} ' or f'\n{tc_id}\n' in text:
            found_page = i
            print(f"Found {tc_id} on page {i+1}")
            break
            
    if found_page != -1:
        page = doc.load_page(found_page)
        pix = page.get_pixmap(dpi=150)
        out_file = os.path.join(out_dir, filename)
        pix.save(out_file)
        print(f"Saved {out_file}")
