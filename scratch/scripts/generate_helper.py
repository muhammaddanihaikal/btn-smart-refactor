import os
import re
import docx

MOBILE_DOCX = r'D:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Mobile.docx'
SCREENSHOT_DIR = r'D:\Project\BTN\Hasil Uji\Screenshot\Mobile'

def sanitize_folder_name(name):
    clean = re.sub(r'[<>:"/\\|?*]', '-', str(name))
    clean = re.sub(r'\s+', ' ', clean).strip(' .')
    if len(clean) > 100:
        clean = clean[:100].strip(' .')
    return clean

doc_src = docx.Document(MOBILE_DOCX)
tcs = []
current_mod = ''
mod_index = 0
last_mod = ''

for child in doc_src.element.body:
    if child.tag.endswith('p'):
        p = docx.text.paragraph.Paragraph(child, doc_src)
        txt = p.text.strip()
        if 'Modul ' in txt and txt[0].isdigit():
            current_mod = txt.split('Modul ')[-1].strip()
            if current_mod != last_mod:
                mod_index += 1
                last_mod = current_mod
    elif child.tag.endswith('tbl'):
        tbl = docx.table.Table(child, doc_src)
        if len(tbl.columns) == 7:
            for r in tbl.rows:
                c0 = r.cells[0].text.strip()
                if c0 and '.' in c0 and c0[0].isdigit():
                    tcs.append({
                        'mod_idx': mod_index,
                        'mod_name': current_mod,
                        'no': c0,
                        'title': r.cells[1].text.strip(),
                        'scenario': r.cells[2].text.strip(),
                        'expected': r.cells[3].text.strip()
                    })

for tc in tcs:
    folder_sub = f"{tc['mod_idx']:02d}. {sanitize_folder_name(tc['mod_name'])}"
    folder_tc = f"{tc['no']} {sanitize_folder_name(tc['title'])}"
    target_dir = os.path.join(SCREENSHOT_DIR, folder_sub, folder_tc)
    
    os.makedirs(target_dir, exist_ok=True)
    
    txt_path = os.path.join(target_dir, f"{folder_tc}.txt")
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(f'''NO TEST CASE : {tc['no']}
JUDUL        : {tc['title']}
MODUL        : {tc['mod_name']}

==================================================
SCENARIO / STEPS:
==================================================
{tc['scenario']}

==================================================
EXPECTED RESULTS (HASIL YANG DIHARAPKAN):
==================================================
{tc['expected']}
''')

print('Generated 187 Helper TXT files!')
