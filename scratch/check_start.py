import zipfile
import xml.etree.ElementTree as ET

docx_path = r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx'
with zipfile.ZipFile(docx_path, 'r') as z:
    doc_xml = z.read('word/document.xml').decode('utf-8')

root = ET.fromstring(doc_xml)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

body = root.find(f'{{{W}}}body')
children = list(body)

for idx in range(min(10, len(children))):
    c = children[idx]
    tag = c.tag.split('}')[-1]
    text = ''.join(c.itertext()).strip()[:60]
    xml_str = ET.tostring(c, encoding='unicode')
    has_sectPr = 'sectPr' in xml_str
    print(f'child {idx}: tag={tag}, has_sectPr={has_sectPr}, text="{text}"')
