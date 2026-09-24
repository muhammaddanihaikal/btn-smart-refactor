import docx
from docx.oxml.ns import qn
import xml.etree.ElementTree as ET

doc = docx.Document(r'd:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx')
body = doc.element.body
for idx, c in enumerate(body):
    tag = c.tag.split('}')[-1]
    xml_str = ET.tostring(c, encoding='unicode')
    has_pbb = 'pageBreakBefore' in xml_str
    has_br = 'w:type="page"' in xml_str or "w:type='page'" in xml_str
    has_sectPr = 'sectPr' in xml_str
    text = ''.join(c.itertext()).strip()[:40]
    print(f'child {idx}: {tag} | sectPr={has_sectPr} | pbb={has_pbb} | br={has_br} | text="{text}"')
