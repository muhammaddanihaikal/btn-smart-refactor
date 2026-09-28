import docx

doc = docx.Document(r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx')

current_module = ""
tc_list = []

for tbl in doc.tables:
    r0 = tbl.rows[0].cells[0].text.strip()
    if r0 == 'No.':
        no = tbl.rows[1].cells[0].text.strip()
        title = tbl.rows[1].cells[1].text.strip().replace('\n', ' ')
        exp = tbl.rows[3].cells[1].text.strip().replace('\n', ' ') if len(tbl.rows) > 3 else ''
        tc_list.append((no, title, exp, True))
    elif '.' in r0 and any(c.isdigit() for c in r0):
        no = r0
        title = tbl.rows[0].cells[1].text.strip().replace('\n', ' ')
        exp = tbl.rows[2].cells[1].text.strip().replace('\n', ' ') if len(tbl.rows) > 2 else ''
        tc_list.append((no, title, exp, False))

print(f"Total TCs: {len(tc_list)}")

# Analyze redundancy
print("\n" + "="*80)
print("ANALISIS DETAIL TC BERMASALAH / REDUNDAN DI MOBILE")
print("="*80)

for i, (no, title, exp, is_first) in enumerate(tc_list):
    prev_no, prev_title, prev_exp, _ = tc_list[i-1] if i > 0 else ("", "", "", False)
    
    # 1. Kasus Sangat Jelas (100% Redundan): TC "Menampilkan daftar/data" tepat setelah "Membuka halaman"
    if 'membuka' in prev_title.lower() and ('menampilkan daftar' in title.lower() or 'menampilkan data' in title.lower() or 'melihat daftar' in title.lower()):
        print(f"\n[KATEGORI 1: 100% REDUNDAN LANGSUNG SETELAH 'MEMBUKA HALAMAN']")
        print(f"  TC Bermasalah : {no} - {title}")
        print(f"  TC Sebelumnya : {prev_no} - {prev_title}")
        print(f"  Expected {prev_no} : {prev_exp}")
        print(f"  Expected {no}   : {exp}")
        print(f"  -> Alasan: Membuka halaman otomatis menampilkan data/daftarnya. Kedua TC punya aksi & SS yang sama.")

    # 2. Kasus Tab Default: TC "Melihat daftar tab X" padahal tab X adalah tab bawaan saat membuka halaman
    elif 'membuka' in prev_title.lower() and 'tab' in title.lower() and any(w in title.lower() for w in ['melihat', 'menampilkan']):
        print(f"\n[KATEGORI 2: REDUNDAN KARENA MERUPAKAN TAB DEFAULT / HALAMAN AWAL]")
        print(f"  TC Bermasalah : {no} - {title}")
        print(f"  TC Sebelumnya : {prev_no} - {prev_title}")
        print(f"  Expected {prev_no} : {prev_exp}")
        print(f"  Expected {no}   : {exp}")
        print(f"  -> Alasan: Saat halaman dibuka ({prev_no}), tab ini adalah tab landing default. Tidak ada aksi klik tab.")

    # 3. Kasus Tab Berturut-turut yang hanya menampilkan list data tanpa aksi bisnis (Aturan Tab di Obsidian: Cukup 1 TC Navigasi Tab)
    elif 'tab' in title.lower() and any(w in title.lower() for w in ['melihat daftar', 'menampilkan data']):
        print(f"\n[KATEGORI 3: TC TAB HANYA UNTUK MELIHAT DAFTAR DATA]")
        print(f"  TC : {no} - {title}")
        print(f"  Expected: {exp}")

    # 4. Kasus TC mandiri "Menampilkan daftar" tanpa aksi (seperti 17.3)
    elif 'menampilkan daftar' in title.lower() or 'melihat daftar' in title.lower():
        print(f"\n[KATEGORI 4: 'MENAMPILKAN / MELIHAT DAFTAR' LAINNYA]")
        print(f"  TC : {no} - {title}")
        print(f"  Expected: {exp}")
