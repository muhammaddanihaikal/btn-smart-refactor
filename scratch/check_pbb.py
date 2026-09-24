import zipfile
import xml.etree.ElementTree as ET

docx_path = r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx'
with zipfile.ZipFile(docx_path, 'r') as z:
    doc_xml = z.read('word/document.xml').decode('utf-8')

root = ET.fromstring(doc_xml)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

body = root.find(f'{{{W}}}body')
children = list(body)

pbb_elements = []
for idx, c in enumerate(children):
    xml_str = ET.tostring(c, encoding='unicode')
    if 'pageBreakBefore' in xml_str:
        tag = c.tag.split('}')[-1]
        text = ''.join(c.itertext()).strip()[:40]
        pbb_elements.append((idx, tag, text, xml_str.count('pageBreakBefore')))

print(f'Total elements with pageBreakBefore: {len(pbb_elements)}')
for idx, tag, text, cnt in pbb_elements[:15]:
    print(f'child {idx}: tag={tag}, cnt={cnt}, text="{text}"')
if len(pbb_elements) > 15:
    print('...')
    for idx, tag, text, cnt in pbb_elements[-5:]:
        print(f'child {idx}: tag={tag}, cnt={cnt}, text="{text}"')
