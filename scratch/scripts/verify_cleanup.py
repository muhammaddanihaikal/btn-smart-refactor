import os, docx, openpyxl

print("=== 1. VERIFY SCREENSHOT FOLDERS (LOCAL D) ===")
base = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'
for mod in ['09. Cuti & Izin', '14. Agenda - Pengingat', '15. Agenda - Daily Sales Agenda']:
    p = os.path.join(base, mod)
    print(f"\nFolder: {mod}")
    for sub in sorted(os.listdir(p)):
        print(f"  {sub}")

print("\n=== 2. VERIFY SIT MOBILE DOCX (LOCAL D) ===")
doc_sit = docx.Document(r'D:\Project\BTN Smart\Refactor\SIT\SIT BTN SMART Mobile.docx')
for t_id, mod_name in [(10, 'Modul 09'), (15, 'Modul 14'), (16, 'Modul 15')]:
    t = doc_sit.tables[t_id]
    print(f"\n{mod_name} in SIT:")
    for r in t.rows[1:]:
        print(f"  {r.cells[0].text:5s} | {r.cells[1].text}")

print("\n=== 3. VERIFY DOKUMEN HASIL UJI MOBILE (LOCAL D) ===")
doc_hu = docx.Document(r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx')
print(f"Total tables in Hasil Uji: {len(doc_hu.tables)}")
for t in doc_hu.tables:
    r0 = t.rows[0].cells[0].text.strip()
    tc = ''
    title = ''
    if r0 == 'No.':
        tc = t.rows[1].cells[0].text.strip()
        title = t.rows[1].cells[1].text.strip().replace('\n', ' ')
    elif '.' in r0 and any(c.isdigit() for c in r0):
        tc = r0
        title = t.rows[0].cells[1].text.strip().replace('\n', ' ')
    if any(tc.startswith(p) for p in ['9.', '14.', '15.']):
        print(f"  {tc:5s} | {title}")

print("\n=== 4. VERIFY SIZES ACROSS DRIVES ===")
files = [
    r'Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx',
    r'SIT\SIT BTN SMART Mobile.docx',
    r'Test Script\Uji Sistem.xlsx'
]
drives = [
    ('Local D', r'D:\Project\BTN Smart\Refactor'),
    ('Drive H', r'H:\My Drive\Zegen\BTN Smart\Refactor'),
    ('Drive G', r'G:\My Drive\Zegen\BTN Smart\Refactor')
]
for f in files:
    print(f"\nFile: {f}")
    for d_name, d_path in drives:
        fp = os.path.join(d_path, f)
        sz = os.path.getsize(fp) if os.path.exists(fp) else 'MISSING'
        print(f"  {d_name:10s}: {sz}")
