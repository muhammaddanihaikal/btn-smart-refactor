import docx

doc = docx.Document(r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx')

all_tcs = []
for i, tbl in enumerate(doc.tables):
    r0_c0 = tbl.rows[0].cells[0].text.strip()
    if r0_c0 == 'No.':
        tc_num = tbl.rows[1].cells[0].text.strip()
        tc_title = tbl.rows[1].cells[1].text.strip().replace('\n', ' ')
        all_tcs.append((i, tc_num, tc_title))
    elif '.' in r0_c0 and any(c.isdigit() for c in r0_c0):
        tc_num = r0_c0
        tc_title = tbl.rows[0].cells[1].text.strip().replace('\n', ' ')
        all_tcs.append((i, tc_num, tc_title))

print(f"Total TCs found in Mobile docx: {len(all_tcs)}")

keywords = [
    'menampilkan daftar',
    'menampilkan data',
    'menampilkan list',
    'melihat daftar',
    'melihat data',
    'menampilkan tabel',
    'melihat tabel'
]

suspects = []
for idx, (tbl_id, tc_num, tc_title) in enumerate(all_tcs):
    prev_tc = all_tcs[idx-1] if idx > 0 else (None, '', '')
    title_lower = tc_title.lower()
    
    is_suspect = False
    matched_kw = ''
    for kw in keywords:
        if kw in title_lower:
            is_suspect = True
            matched_kw = kw
            break
            
    if is_suspect:
        suspects.append((tbl_id, tc_num, tc_title, prev_tc[1], prev_tc[2], matched_kw))

print(f"\n=== CANDIDATE TCs WITH 'MENAMPILKAN / MELIHAT DAFTAR/DATA' ({len(suspects)}) ===\n")
for tbl_id, tc_num, tc_title, prev_num, prev_title, kw in suspects:
    print(f"TC {tc_num} (Table {tbl_id:3d}): \"{tc_title}\" [Keyword: '{kw}']")
    print(f"   Preceded by TC {prev_num}: \"{prev_title}\"")
    if 'membuka' in prev_title.lower():
        print(f"   >>> REDUNDANT ALERT: Follows 'Membuka' directly!")
    print()
