import pymupdf as fitz
doc = fitz.open(r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf')
print("Page 34 text:")
print(doc[33].get_text()[:400])
