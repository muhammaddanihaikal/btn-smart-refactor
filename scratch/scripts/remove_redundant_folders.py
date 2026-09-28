import os, shutil, re

def update_txt_file(folder_path, new_no, new_title):
    # Find txt file in folder
    for f in os.listdir(folder_path):
        if f.endswith('.txt'):
            old_txt = os.path.join(folder_path, f)
            with open(old_txt, 'r', encoding='utf-8', errors='ignore') as tf:
                content = tf.read()
            
            # replace NO TEST CASE
            content = re.sub(r'NO TEST CASE\s*:\s*[^\n]+', f'NO TEST CASE : {new_no}', content)
            # replace JUDUL if needed
            if new_title:
                content = re.sub(r'JUDUL\s*:\s*[^\n]+', f'JUDUL        : {new_title}', content)
                
            os.remove(old_txt)
            new_txt_name = f"{new_no} {new_title}.txt" if new_title else f"{new_no}.txt"
            new_txt_path = os.path.join(folder_path, new_txt_name)
            with open(new_txt_path, 'w', encoding='utf-8') as tf:
                tf.write(content)
            # print(f"    Updated txt: {new_txt_name}")
            break

def process_base(screenshot_base):
    print(f"\n==========================================")
    print(f"Processing Screenshot Base: {screenshot_base}")
    print(f"==========================================")
    if not os.path.exists(screenshot_base):
        print(f"Base path does not exist: {screenshot_base}")
        return

    # 1. Modul 09. Cuti & Izin
    mod_09 = os.path.join(screenshot_base, '09. Cuti & Izin')
    if os.path.exists(mod_09):
        print(f"\n--- Modul 09. Cuti & Izin ---")
        # Folder to delete
        del_target = os.path.join(mod_09, '9.2 Menampilkan daftar data pengajuan cuti')
        if os.path.exists(del_target):
            shutil.rmtree(del_target)
            print(f"Deleted redundant folder: 9.2 Menampilkan daftar data pengajuan cuti")
            
        # Shifts: 9.3 -> 9.2, 9.4 -> 9.3, ..., 9.9 -> 9.8
        shifts_09 = [
            ('9.3', '9.2', 'Mencari data pengajuan cuti dengan keyword valid'),
            ('9.4', '9.3', 'Mencari data pengajuan cuti dengan keyword tidak valid'),
            ('9.5', '9.4', 'Melakukan filter data pengajuan cuti'),
            ('9.6', '9.5', 'Menampilkan data pengajuan cuti berdasarkan jenis cuti'),
            ('9.7', '9.6', 'Menampilkan data pengajuan cuti berdasarkan status'),
            ('9.8', '9.7', 'Mengajukan cuti'),
            ('9.9', '9.8', 'Mengajukan cuti tanpa mengisi field mandatory'),
        ]
        
        # Phase 1: Rename to temp
        for old_no, new_no, title in shifts_09:
            old_name = f"{old_no} {title}"
            old_path = os.path.join(mod_09, old_name)
            tmp_path = os.path.join(mod_09, f"__tmp_09_{new_no}")
            if os.path.exists(old_path):
                os.rename(old_path, tmp_path)
                
        # Phase 2: Rename from temp to new
        for old_no, new_no, title in shifts_09:
            tmp_path = os.path.join(mod_09, f"__tmp_09_{new_no}")
            new_name = f"{new_no} {title}"
            new_path = os.path.join(mod_09, new_name)
            if os.path.exists(tmp_path):
                os.rename(tmp_path, new_path)
                update_txt_file(new_path, new_no, title)
                print(f"Renamed: {old_no} -> {new_no} {title}")

    # 2. Modul 14. Agenda - Pengingat
    mod_14 = os.path.join(screenshot_base, '14. Agenda - Pengingat')
    if os.path.exists(mod_14):
        print(f"\n--- Modul 14. Agenda - Pengingat ---")
        del_target = os.path.join(mod_14, '14.2 Melihat daftar pengingat pada tab Aktif')
        if os.path.exists(del_target):
            shutil.rmtree(del_target)
            print(f"Deleted redundant folder: 14.2 Melihat daftar pengingat pada tab Aktif")
            
        shifts_14 = [
            ('14.3', '14.2', 'Melihat daftar pengingat pada tab Selesai'),
            ('14.4', '14.3', 'Menambahkan pengingat baru dengan data valid'),
            ('14.5', '14.4', 'Menambahkan pengingat baru tanpa mengisi field mandatory'),
            ('14.6', '14.5', 'Menandai pengingat sebagai selesai'),
        ]
        
        # Phase 1: Temp
        for old_no, new_no, title in shifts_14:
            old_name = f"{old_no} {title}"
            old_path = os.path.join(mod_14, old_name)
            tmp_path = os.path.join(mod_14, f"__tmp_14_{new_no}")
            if os.path.exists(old_path):
                os.rename(old_path, tmp_path)
                
        # Phase 2: Final
        for old_no, new_no, title in shifts_14:
            tmp_path = os.path.join(mod_14, f"__tmp_14_{new_no}")
            new_name = f"{new_no} {title}"
            new_path = os.path.join(mod_14, new_name)
            if os.path.exists(tmp_path):
                os.rename(tmp_path, new_path)
                update_txt_file(new_path, new_no, title)
                print(f"Renamed: {old_no} -> {new_no} {title}")

    # 3. Modul 15. Agenda - Daily Sales Agenda
    mod_15 = os.path.join(screenshot_base, '15. Agenda - Daily Sales Agenda')
    if os.path.exists(mod_15):
        print(f"\n--- Modul 15. Agenda - Daily Sales Agenda ---")
        del_target = os.path.join(mod_15, '15.2 Melihat daftar agenda pada tab Masuk')
        if os.path.exists(del_target):
            shutil.rmtree(del_target)
            print(f"Deleted redundant folder: 15.2 Melihat daftar agenda pada tab Masuk")
            
        shifts_15 = [
            ('15.3', '15.2', 'Melihat daftar agenda pada tab Closing'),
            ('15.4', '15.3', 'Melihat daftar agenda pada tab Referral'),
            ('15.5', '15.4', 'Melihat daftar agenda pada tab Ref. Closing'),
            ('15.6', '15.5', 'Mencari agenda menggunakan Search Bar'),
            ('15.7', '15.6', 'Mencari agenda dengan keyword tidak valid'),
            ('15.8', '15.7', 'Melakukan filter data Daily Sales Agenda'),
        ]
        
        # Phase 1: Temp
        for old_no, new_no, title in shifts_15:
            old_name = f"{old_no} {title}"
            old_path = os.path.join(mod_15, old_name)
            tmp_path = os.path.join(mod_15, f"__tmp_15_{new_no}")
            if os.path.exists(old_path):
                os.rename(old_path, tmp_path)
                
        # Phase 2: Final
        for old_no, new_no, title in shifts_15:
            tmp_path = os.path.join(mod_15, f"__tmp_15_{new_no}")
            new_name = f"{new_no} {title}"
            new_path = os.path.join(mod_15, new_name)
            if os.path.exists(tmp_path):
                os.rename(tmp_path, new_path)
                update_txt_file(new_path, new_no, title)
                print(f"Renamed: {old_no} -> {new_no} {title}")

# Execute across all 3 bases
bases = [
    r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile',
    r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile',
    r'G:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'
]

for b in bases:
    process_base(b)

print("\nAll screenshot folder renames & txt updates finished successfully!")
