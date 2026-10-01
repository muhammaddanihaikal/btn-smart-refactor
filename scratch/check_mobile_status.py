import docx
import xml.etree.ElementTree as ET
from collections import defaultdict

doc = docx.Document(r'D:/Project/BTN Smart/Refactor/Hasil Uji/Dokumen_Hasil_Uji_Mobile.docx')
module_stats = defaultdict(lambda: {'total': 0, 'filled': 0, 'missing': []})

for i, t in enumerate(doc.tables):
    if len(t.rows) < 3 or len(t.rows[0].cells) < 2:
        continue
    tc_num = ''
    tc_title = ''
    img_row = 1
    
    if t.rows[0].cells[0].text.strip() == 'No.':
        tc_num = t.rows[1].cells[0].text.strip()
        tc_title = t.rows[1].cells[1].text.strip()
        img_row = 2
    else:
        tc_num = t.rows[0].cells[0].text.strip()
        tc_title = t.rows[0].cells[1].text.strip()
        img_row = 1
        
    if not tc_num or tc_title.startswith('User Acceptance') or 'Informasi Dokumen' in tc_title or 'Pelaksana' in tc_title:
        continue
        
    mod_id = tc_num.split('.')[0]
    
    has_drawing = False
    if img_row < len(t.rows):
        for cell in t.rows[img_row].cells:
            root = ET.fromstring(cell._tc.xml)
            for elem in root.iter():
                if elem.tag.endswith(('drawing', 'pict')):
                    has_drawing = True
                    break
            if has_drawing:
                break
                
    module_stats[mod_id]['total'] += 1
    if has_drawing:
        module_stats[mod_id]['filled'] += 1
    else:
        module_stats[mod_id]['missing'].append((tc_num, tc_title))

for mod in sorted(module_stats.keys(), key=lambda x: int(x) if x.isdigit() else 99):
    st = module_stats[mod]
    pct = int(st['filled'] / st['total'] * 100) if st['total'] > 0 else 0
    status = 'LENGKAP (100%)' if st['total'] == st['filled'] else f"{st['filled']}/{st['total']} ({pct}%)"
    print(f'Modul {mod:2s}: {status}')
    if st['missing']:
        for num, title in st['missing']:
            print(f'   - {num}: {title}')
