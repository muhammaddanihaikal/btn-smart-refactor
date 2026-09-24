import pymupdf as fitz
doc = fitz.open(r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf')
print("--- PAGE 5 ---")
print(doc[4].get_text())
print("--- PAGE 6 ---")
print(doc[5].get_text())
