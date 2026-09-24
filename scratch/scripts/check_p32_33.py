import pymupdf as fitz
doc = fitz.open(r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf')
print("Page 32 text:")
print(doc[31].get_text()[:400])
print("\nPage 33 text:")
print(doc[32].get_text()[:400])
