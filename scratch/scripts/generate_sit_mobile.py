import docx, re, copy, zipfile, shutil, os
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from collections import OrderedDict

# ── PATHS ─────────────────────────────────────────────────────────────────
SENIOR_DOCX   = r'D:\Project\BTN\SIT\SIT BTN SMART Mobile (Updated).docx'
TEMPLATE_PATH = r'D:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx'
OUTPUT_DIR_LOCAL   = r'D:\Project\BTN\SIT\Document SIT BTN Smart'
OUTPUT_MOBILE_LOCAL = os.path.join(OUTPUT_DIR_LOCAL, 'SIT BTN SMART Mobile.docx')
OUTPUT_MOBILE_H = r'H:\My Drive\Zegen\BTN Smart\Refactor\SIT\SIT BTN SMART Mobile.docx'
OUTPUT_MOBILE_G = r'G:\My Drive\Zegen\BTN Smart\Refactor\SIT\SIT BTN SMART Mobile.docx'
os.makedirs(OUTPUT_DIR_LOCAL, exist_ok=True)

# ── 1. LOAD TC DATA FROM SENIOR DOCX ─────────────────────────────────────
print('Loading TCs from: ' + SENIOR_DOCX)
senior_doc = docx.Document(SENIOR_DOCX)

raw_module_map = {}  # m_num -> list of tc dicts (raw from senior doc)
for tbl in senior_doc.tables:
    if len(tbl.columns) == 7:
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            no = cells[0]
            if no.lower() in ('no', 'no.'): continue
            if not no or '.' not in no: continue
            try:
                m_num = int(no.split('.')[0])
            except ValueError:
                continue
            if m_num not in raw_module_map:
                raw_module_map[m_num] = []
            raw_module_map[m_num].append({
                'no': no,
                'title': cells[1],
                'scenario': cells[2],
                'expected': cells[3],
            })

# ── 2. BUILD GRANULAR SUBMENU LIST ─────────────────────────────────────────
# Split logic based on TC title boundaries confirmed from Uji Sistem.xlsx:
#   Mod 3: 3.1-3.6 = Visit In, 3.7-3.12 = Visit Out
#   Mod 4: 4.1-4.7 = Clock In, 4.8-4.14 = Clock Out
#   Mod 8: 8.1-8.21 = Aktifitas Prospek, 8.22-8.41 = Input Prospek,
#           8.42-8.51 = List Nasabah ETB, 8.52-8.55 = Nasabah Referal
#   Mod 9: 9.1-9.6 = Pengingat, 9.7-9.14 = Daily Sales Agenda

def split_by_no(tcs, boundary_no):
    """Split list into two at the given TC number (boundary goes to second half)."""
    bound_idx = int(boundary_no.split('.')[1])
    before = [t for t in tcs if int(t['no'].split('.')[1]) < bound_idx]
    after  = [t for t in tcs if int(t['no'].split('.')[1]) >= bound_idx]
    return before, after

def split_three(tcs, b1, b2):
    a, rest = split_by_no(tcs, b1)
    b, c    = split_by_no(rest, b2)
    return a, b, c

def split_four(tcs, b1, b2, b3):
    a, rest1 = split_by_no(tcs, b1)
    b, rest2 = split_by_no(rest1, b2)
    c, d     = split_by_no(rest2, b3)
    return a, b, c, d

# OrderedDict: (mod_name, sub_name) -> [tc_dicts]
final_submenus = OrderedDict()

for m_num in sorted(raw_module_map.keys()):
    tcs = raw_module_map[m_num]

    if m_num == 1:
        final_submenus[('Login', '')] = tcs
    elif m_num == 2:
        final_submenus[('Profile', '')] = tcs
    elif m_num == 3:
        visit_in, visit_out = split_by_no(tcs, '3.7')
        final_submenus[('Sales Tracking Activity', 'Visit In')]  = visit_in
        final_submenus[('Sales Tracking Activity', 'Visit Out')] = visit_out
    elif m_num == 4:
        clock_in, clock_out = split_by_no(tcs, '4.8')
        final_submenus[('Absent', 'Clock In')]  = clock_in
        final_submenus[('Absent', 'Clock Out')] = clock_out
    elif m_num == 5:
        final_submenus[('Log Absensi', '')] = tcs
    elif m_num == 6:
        final_submenus[('Personal Funnel', '')] = tcs
    elif m_num == 7:
        final_submenus[('Cuti & Izin', '')] = tcs
    elif m_num == 8:
        # 8.1-8.21=Aktifitas Prospek, 8.22-8.41=Input Prospek,
        # 8.42-8.51=List Nasabah ETB, 8.52-8.55=Nasabah Referal
        aktifitas, input_p, list_etb, referal = split_four(tcs, '8.22', '8.42', '8.52')
        final_submenus[('Prospek & Nasabah', 'Aktifitas Prospek')]  = aktifitas
        final_submenus[('Prospek & Nasabah', 'Input Prospek')]      = input_p
        final_submenus[('Prospek & Nasabah', 'List Nasabah ETB')]   = list_etb
        final_submenus[('Prospek & Nasabah', 'Nasabah Referal')]    = referal
    elif m_num == 9:
        # 9.1-9.6=Pengingat, 9.7-9.14=Daily Sales Agenda
        pengingat, daily = split_by_no(tcs, '9.7')
        final_submenus[('Agenda', 'Pengingat')]           = pengingat
        final_submenus[('Agenda', 'Daily Sales Agenda')]  = daily
    elif m_num == 10:
        final_submenus[('Fitur', 'Marketing Toolkit')] = tcs
    elif m_num == 11:
        final_submenus[('Sales Force', 'Dashboard Sales Code')] = tcs

total_tcs = sum(len(v) for v in final_submenus.values())
print('Total sections : ' + str(len(final_submenus)))
print('Total TCs      : ' + str(total_tcs))
for idx, ((mod, sub), tcs) in enumerate(final_submenus.items(), 1):
    label = mod + (' - ' + sub if sub else '')
    print('  ' + str(idx) + '. Modul ' + label + ': ' + str(len(tcs)) + ' TC')

# ── 3. HELPERS ─────────────────────────────────────────────────────────────
def clean_no_numbers(text):
    if not text: return ''
    cleaned = []
    for ln in text.split('\n'):
        ln_c = re.sub(r'^\s*\d+[\.)\-]\s*', '', ln.strip())
        if ln_c: cleaned.append(ln_c)
    return '\n'.join(cleaned)

def make_clean_actual(expected_text):
    clean = clean_no_numbers(expected_text)
    result = []
    for l in clean.split('\n'):
        l_s = l.strip()
        if l_s and not l_s.lower().startswith('berhasil'):
            l_s = 'Berhasil ' + l_s
        result.append(l_s)
    return '\n'.join(result)

# ── 4. XML HELPERS ─────────────────────────────────────────────────────────
doc_template    = Document(TEMPLATE_PATH)
body_tpl        = list(doc_template.element.body)
TBL_HEADER      = doc_template.tables[0]
TBL_INFO        = doc_template.tables[1]
TBL_DATA_TPL    = doc_template.tables[2]
DAFTAR_TEMPLATE = body_tpl[20]

def set_cell_center(tc_elem, text, bold=False):
    tcPr = tc_elem.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr'); tc_elem.insert(0, tcPr)
    va = tcPr.find(qn('w:vAlign'))
    if va is None:
        va = OxmlElement('w:vAlign'); tcPr.append(va)
    va.set(qn('w:val'), 'center')

    text = text if isinstance(text, str) else (str(text) if text else '')
    lines = text.split('\n')
    paras = tc_elem.findall(qn('w:p'))
    while len(tc_elem.findall(qn('w:p'))) < len(lines):
        p_new = copy.deepcopy(paras[0])
        for ch in list(p_new):
            if ch.tag != qn('w:pPr'): p_new.remove(ch)
        tc_elem.append(p_new)

    paras = tc_elem.findall(qn('w:p'))
    for i, line in enumerate(lines):
        p = paras[i]
        pPr = p.find(qn('w:pPr'))
        if pPr is None: pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
        jc = pPr.find(qn('w:jc'))
        if jc is None: jc = OxmlElement('w:jc'); pPr.append(jc)
        jc.set(qn('w:val'), 'center')
        for ch in list(p):
            if ch.tag not in [qn('w:pPr'), qn('w:pPrChange')]: p.remove(ch)
        r_elem = OxmlElement('w:r')
        if bold:
            rPr = OxmlElement('w:rPr'); rPr.append(OxmlElement('w:b')); r_elem.append(rPr)
        t = OxmlElement('w:t')
        t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r_elem.append(t); p.append(r_elem)
    for extra in tc_elem.findall(qn('w:p'))[len(lines):]:
        tc_elem.remove(extra)

def build_info_table(submenus_list, total_tc):
    tbl = copy.deepcopy(TBL_INFO._tbl)
    rows = tbl.findall(qn('w:tr'))
    r0_cells = rows[0].findall(qn('w:tc'))
    set_cell_center(r0_cells[1], 'Jumlah Script : ' + str(total_tc))

    r1c0 = rows[1].findall(qn('w:tc'))[0]
    orig_paras = r1c0.findall(qn('w:p'))
    bullet_pPr = None
    if len(orig_paras) > 1:
        pPr = orig_paras[1].find(qn('w:pPr'))
        if pPr is not None: bullet_pPr = copy.deepcopy(pPr)

    for p in r1c0.findall(qn('w:p')):
        for ch in list(p):
            if ch.tag not in [qn('w:pPr'), qn('w:pPrChange')]: p.remove(ch)

    needed = 1 + len(submenus_list)
    while len(r1c0.findall(qn('w:p'))) < needed:
        r1c0.append(OxmlElement('w:p'))

    paras = r1c0.findall(qn('w:p'))
    p0 = paras[0]
    r0 = OxmlElement('w:r')
    r0Pr = OxmlElement('w:rPr'); r0Pr.append(OxmlElement('w:b'))
    sz0 = OxmlElement('w:sz'); sz0.set(qn('w:val'), '20'); r0Pr.append(sz0)
    r0.append(r0Pr)
    t0 = OxmlElement('w:t'); t0.text = 'Modul :'; r0.append(t0); p0.append(r0)

    for idx, (mod_name, sub_name) in enumerate(submenus_list):
        label = 'Modul ' + mod_name + (' - ' + sub_name if sub_name else '')
        p = paras[1 + idx]
        if bullet_pPr is not None:
            old = p.find(qn('w:pPr'))
            if old is not None: p.remove(old)
            cloned = copy.deepcopy(bullet_pPr)
            sp = cloned.find(qn('w:spacing'))
            if sp is not None: sp.set(qn('w:before'), '20'); sp.set(qn('w:after'), '0')
            p.insert(0, cloned)
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '20'); rPr.append(sz); r.append(rPr)
        t = OxmlElement('w:t'); t.text = label
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t); p.append(r)

    set_cell_center(rows[2].findall(qn('w:tc'))[0], '')
    return tbl

def build_title_para(sec_idx, mod_name, sub_name):
    p = copy.deepcopy(DAFTAR_TEMPLATE)
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        for tag in [qn('w:numPr'), qn('w:tabs')]:
            el = pPr.find(tag)
            if el is not None: pPr.remove(el)
        ind = pPr.find(qn('w:ind'))
        if ind is None: ind = OxmlElement('w:ind'); pPr.append(ind)
        ind.set(qn('w:left'), '259'); ind.set(qn('w:hanging'), '0')
        sp = pPr.find(qn('w:spacing'))
        if sp is None: sp = OxmlElement('w:spacing'); pPr.append(sp)
        sp.set(qn('w:before'), '120'); sp.set(qn('w:after'), '160')
        if sec_idx > 1:
            pbb = pPr.find(qn('w:pageBreakBefore'))
            if pbb is None:
                pbb = OxmlElement('w:pageBreakBefore')
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None: pStyle.addnext(pbb)
                else: pPr.insert(0, pbb)

    orig_runs = p.findall(qn('w:r'))
    orig_rpr = None
    if orig_runs:
        rPr_el = orig_runs[0].find(qn('w:rPr'))
        if rPr_el is not None: orig_rpr = copy.deepcopy(rPr_el)
    for ch in list(p):
        if ch.tag not in [qn('w:pPr'), qn('w:pPrChange')]: p.remove(ch)

    label = 'Modul ' + mod_name + (' - ' + sub_name if sub_name else '')
    title = str(sec_idx) + '. ' + label
    r_elem = OxmlElement('w:r')
    if orig_rpr is not None: r_elem.append(orig_rpr)
    t = OxmlElement('w:t'); t.text = title
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r_elem.append(t); p.append(r_elem)
    return p

def zero_height_para(sect_pr=None):
    p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr')
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
    sp.set(qn('w:line'), '1'); sp.set(qn('w:lineRule'), 'exact')
    pPr.append(sp)
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '2'); rPr.append(sz); pPr.append(rPr)
    if sect_pr is not None: pPr.append(copy.deepcopy(sect_pr))
    p.append(pPr); return p

def empty_para():
    p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle'); pStyle.set(qn('w:val'), 'TeksIsi')
    pPr.append(pStyle); p.insert(0, pPr); return p

def build_data_table(tcs, sec_idx):
    tbl = copy.deepcopy(TBL_DATA_TPL._tbl)
    rows_elem = tbl.findall(qn('w:tr'))
    for r in rows_elem[1:]: tbl.remove(r)
    header_tr = rows_elem[0]
    for c in header_tr.findall(qn('w:tc')):
        tcPr = c.find(qn('w:tcPr'))
        if tcPr is not None:
            va = tcPr.find(qn('w:vAlign'))
            if va is None: va = OxmlElement('w:vAlign'); tcPr.append(va)
            va.set(qn('w:val'), 'center')
        for p in c.findall(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is None: pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
            jc = pPr.find(qn('w:jc'))
            if jc is None: jc = OxmlElement('w:jc'); pPr.append(jc)
            jc.set(qn('w:val'), 'center')

    tmpl_row = TBL_DATA_TPL.rows[1]
    for i, tc in enumerate(tcs, start=1):
        tc_number = str(sec_idx) + '.' + str(i)
        clean_scenario = clean_no_numbers(tc['scenario'])
        clean_exp = clean_no_numbers(tc['expected'])
        clean_act = make_clean_actual(clean_exp)
        new_tr = copy.deepcopy(tmpl_row._tr)
        cells = new_tr.findall(qn('w:tc'))
        set_cell_center(cells[0], tc_number)
        set_cell_center(cells[1], tc['title'])
        set_cell_center(cells[2], clean_scenario)
        set_cell_center(cells[3], clean_exp)
        set_cell_center(cells[4], clean_act)
        set_cell_center(cells[5], 'P')
        set_cell_center(cells[6], '')
        tbl.append(new_tr)
    return tbl

# ── 5. POST-PROCESS (header + PAGE field) ────────────────────────────────

def strip_para_ids(docx_path):
    tmp = docx_path + '.tmp_para'
    with zipfile.ZipFile(docx_path, 'r') as z_in, zipfile.ZipFile(tmp, 'w') as z_out:
        for item in z_in.infolist():
            if item.filename == 'word/document.xml':
                content = z_in.read(item.filename)
                try:
                    import lxml.etree as etree
                    root = etree.fromstring(content)
                    for el in root.iter():
                        for k in list(el.attrib.keys()):
                            if 'paraId' in k or 'textId' in k:
                                del el.attrib[k]
                    content = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
                except: pass
                z_out.writestr(item, content)
            else:
                z_out.writestr(item, z_in.read(item.filename))
    import shutil
    shutil.move(tmp, docx_path)

def post_process(docx_path):
    tmp = docx_path + '.tmp_pp'
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as z_src:
        template_header = (z_src.read('word/header1.xml').decode('utf-8')
                           if 'word/header1.xml' in z_src.namelist() else None)
    if template_header:
        page_idx = template_header.find(' PAGE </w:instrText>')
        if page_idx != -1:
            ac_start = template_header.rfind('<mc:AlternateContent>', 0, page_idx)
            ac_end = template_header.find('</mc:AlternateContent>', page_idx) + len('</mc:AlternateContent>')
            r_start = template_header.rfind('<w:r ', 0, ac_start)
            r_end = template_header.find('</w:r>', ac_end) + len('</w:r>')
            template_header = template_header[:r_start] + template_header[r_end:]
            inline_page = (
                '<w:p><w:pPr><w:pStyle w:val="TeksIsi"/><w:jc w:val="right"/>'
                '<w:spacing w:before="530"/><w:ind w:right="300"/>'
                '<w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>'
                '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>'
                '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
                '<w:t xml:space="preserve">Page </w:t></w:r>'
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>'
                '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
                '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                '<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>'
                '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>'
                '<w:t>2</w:t></w:r>'
                '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            )
            template_header = template_header.replace('</w:hdr>', inline_page + '</w:hdr>')

    with zipfile.ZipFile(docx_path, 'r') as z_in, zipfile.ZipFile(tmp, 'w') as z_out:
        for item in z_in.infolist():
            if item.filename == 'word/header1.xml' and template_header:
                z_out.writestr(item, template_header.encode('utf-8'))
            elif item.filename == 'word/settings.xml':
                content = z_in.read(item.filename)
                try:
                    import lxml.etree as etree
                    root = etree.fromstring(content)
                    for uf in root.findall('.//w:updateFields', namespaces={'w': W}):
                        root.remove(uf)
                    content = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone='yes')
                except Exception:
                    pass
                z_out.writestr(item, content)
            else:
                z_out.writestr(item, z_in.read(item.filename))
    shutil.move(tmp, docx_path)
    print('Post-processed: ' + docx_path)

# ── 6. ASSEMBLE DOCUMENT ──────────────────────────────────────────────────
def generate_sit_mobile(output_path):
    print('\nBuilding: ' + output_path)
    doc = Document(TEMPLATE_PATH)
    body = doc.element.body

    inner_sect_pr = None
    final_sect_pr = None
    for child in list(body):
        if child.tag == qn('w:p'):
            pPr = child.find(qn('w:pPr'))
            if pPr is not None:
                sp = pPr.find(qn('w:sectPr'))
                if sp is not None: inner_sect_pr = copy.deepcopy(sp)
        elif child.tag == qn('w:sectPr'):
            final_sect_pr = copy.deepcopy(child)

    for child in list(body): body.remove(child)

    # 1) Header table
    body.append(copy.deepcopy(TBL_HEADER._tbl))
    body.append(empty_para())

    # 2) Info table (all 13 sub-menu sections)
    body.append(build_info_table(list(final_submenus.keys()), total_tcs))

    # 3) Zero-height section break
    body.append(zero_height_para(inner_sect_pr))

    # 4) 13 sections
    for sec_idx, ((mod_name, sub_name), tcs) in enumerate(final_submenus.items(), 1):
        body.append(build_title_para(sec_idx, mod_name, sub_name))
        body.append(build_data_table(tcs, sec_idx))

    # 5) Trailing para + final sectPr
    body.append(zero_height_para())
    if final_sect_pr is not None:
        pgNum = final_sect_pr.find(qn('w:pgNumType'))
        if pgNum is not None: final_sect_pr.remove(pgNum)
        body.append(final_sect_pr)

    doc.save(output_path)
    print('Saved: ' + output_path)
    post_process(output_path)

# ── 7. RUN ─────────────────────────────────────────────────────────────────
generate_sit_mobile(OUTPUT_MOBILE_LOCAL)
OUTPUT_MOBILE_ROOT = r'D:\Project\BTN\SIT\SIT BTN SMART Mobile.docx'
for dst in [OUTPUT_MOBILE_ROOT, OUTPUT_MOBILE_H, OUTPUT_MOBILE_G]:
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(OUTPUT_MOBILE_LOCAL, dst)
    print('Synced to: ' + dst)
print('\nDone!')
