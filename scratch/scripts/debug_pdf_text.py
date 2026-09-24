import pymupdf as fitz
doc = fitz.open(r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf')
for i in range(len(doc[:10])):
    txt = doc[i].get_text()
    for line in txt.split('\n'):
        if '1.2' in line or '1.3' in line or 'Login' in line:
            print(f"Page {i+1}: {line.strip()}")
