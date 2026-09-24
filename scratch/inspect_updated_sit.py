import docx

doc = docx.Document(r'H:\My Drive\Zegen\BTN Smart\Refactor\SIT\SIT BTN SMART Web (Updated).docx')
body = doc.element.body

cur_heading = 'None'
table_info = []

for child in list(body):
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        t = ''.join(child.itertext()).strip()
        if t and 'Modul' in t:
            cur_heading = t
    elif tag == 'tbl':
        rows = child.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        if rows:
            cells = rows[0].findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            if len(cells) == 7 and len(rows) > 1:
                r1_cells = rows[1].findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
                first_no = ''.join(r1_cells[0].itertext()).strip()
                first_title = ''.join(r1_cells[1].itertext()).strip()
                last_cells = rows[-1].findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
                last_no = ''.join(last_cells[0].itertext()).strip()
                last_title = ''.join(last_cells[1].itertext()).strip()
                table_info.append((cur_heading, len(rows) - 1, first_no, first_title, last_no, last_title))

print(f"Total data tables: {len(table_info)}")
for idx, (head, count, f_no, f_title, l_no, l_title) in enumerate(table_info):
    print(f"Table {idx+1:2d} | [{head}] ({count:2d} TCs): {f_no} '{f_title[:30]}' -> {l_no} '{l_title[:30]}'")
