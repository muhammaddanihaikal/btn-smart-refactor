import pymupdf as fitz
doc = fitz.open(r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf')

page5 = doc.load_page(4) # 0-indexed, so 4 is Page 5
pix5 = page5.get_pixmap(dpi=150)
pix5.save(r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816\tc_1_2_1_3_preview.png')

page6 = doc.load_page(5)
pix6 = page6.get_pixmap(dpi=150)
pix6.save(r'C:\Users\acer\.gemini\antigravity\brain\c1f939b3-9011-4e84-b54d-0f3eb1928816\page_6_preview.png')

print("Pages 5 and 6 saved!")
