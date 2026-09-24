import zipfile
import xml.etree.ElementTree as ET

docx_path = r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx'
with zipfile.ZipFile(docx_path, 'r') as z:
    doc_xml = z.read('word/document.xml').decode('utf-8')

root = ET.fromstring(doc_xml)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

body = root.find(f'{{{W}}}body')
children = list(body)

print("Child 3 XML:")
print(ET.tostring(children[3], encoding='unicode'))
