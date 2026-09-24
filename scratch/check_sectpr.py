import docx
import xml.etree.ElementTree as ET

doc = docx.Document(r'd:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx')
body = doc.element.body

for p in body.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
    pPr = p.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    if pPr is not None:
        sectPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
        if sectPr is not None:
            print("Found inner sectPr:")
            print(ET.tostring(sectPr, encoding='unicode'))

final_sectPr = body.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
if final_sectPr is not None:
    print("Found final sectPr:")
    print(ET.tostring(final_sectPr, encoding='unicode'))
