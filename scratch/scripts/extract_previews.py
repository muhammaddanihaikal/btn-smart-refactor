import pymupdf as fitz
import os

pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'
out_dir = r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816'

doc = fitz.open(pdf_path)

for tc_id, filename in [('1.2', 'tc_1_2_preview.png'), ('1.3', 'tc_1_3_preview.png')]:
    target_page = -1
    for i in range(len(doc)):
        text = doc[i].get_text()
        if f' {tc_id} ' in f' {text} ' or f'\n{tc_id}\n' in text:
            # Check title keywords
            if (tc_id == '1.2' and 'password valid' in text) or (tc_id == '1.3' and 'tidak terdaftar' in text):
                target_page = i
                print(f"Found TC {tc_id} on page {i+1}")
                break
                
    if target_page != -1:
        page = doc.load_page(target_page)
        pix = page.get_pixmap(dpi=150)
        out_path = os.path.join(out_dir, filename)
        pix.save(out_path)
        print(f'Saved preview to {out_path}')
    else:
        print(f"Could not find page for TC {tc_id}")
