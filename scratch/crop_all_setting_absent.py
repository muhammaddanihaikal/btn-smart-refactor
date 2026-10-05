import os
import shutil
import re
from PIL import Image

def win_p(p):
    p = os.path.abspath(p)
    return '\\\\?\\' + p if not p.startswith('\\\\?\\') else p

base = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot\Web'

# Backup current docx
doc_p = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Web.docx'
backup_dir = r'D:\Project\BTN Smart\Refactor\scratch\backups'
os.makedirs(win_p(backup_dir), exist_ok=True)
backup_file = os.path.join(backup_dir, 'Dokumen_Hasil_Uji_Web_before_setting_absent.docx')
shutil.copy2(win_p(doc_p), win_p(backup_file))
print(f"Backed up docx to {backup_file}")

# Module 40 crop definitions
crop_rules = {
    # 40. Setting Absent - Attendance Spot
    # Opening
    ('40. Setting Absent - Attendance Spot', '40.1 Membuka halaman Attendance Spot', '1.png'): (0, 0),
    # Search
    ('40. Setting Absent - Attendance Spot', '40.2 Mencari data Attendance Spot dengan keyword valid', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.3 Mencari data Attendance Spot dengan keyword tidak valid', '1.png'): (338, 72),
    # Add
    ('40. Setting Absent - Attendance Spot', '40.4 Menambahkan data Attendance Spot', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.4 Menambahkan data Attendance Spot', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.4 Menambahkan data Attendance Spot', '3.png'): (338, 0), # Alert
    # Add Mandatory
    ('40. Setting Absent - Attendance Spot', '40.5 Menambahkan data Attendance Spot tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.5 Menambahkan data Attendance Spot tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Map Search
    ('40. Setting Absent - Attendance Spot', '40.6 Mencari lokasi spesifik melalui pencarian peta dengan keyword valid', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.6 Mencari lokasi spesifik melalui pencarian peta dengan keyword valid', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.7 Mencari lokasi spesifik melalui pencarian peta dengan keyword tidak valid', '1.png'): (338, 72),
    # Edit
    ('40. Setting Absent - Attendance Spot', '40.8 Mengubah data Attendance Spot', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.8 Mengubah data Attendance Spot', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.8 Mengubah data Attendance Spot', '3.png'): (338, 0), # Alert
    # Edit Mandatory
    ('40. Setting Absent - Attendance Spot', '40.9 Mengubah data Attendance Spot tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.9 Mengubah data Attendance Spot tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Map Detail Modal
    ('40. Setting Absent - Attendance Spot', '40.10 Melihat detail lokasi pada peta', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.10 Melihat detail lokasi pada peta', '2.png'): (338, 72),
    # Delete
    ('40. Setting Absent - Attendance Spot', '40.11 Menghapus data Attendance Spot', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.11 Menghapus data Attendance Spot', '2.png'): (338, 0), # Alert
    # Setting Personnel
    ('40. Setting Absent - Attendance Spot', '40.12 Membuka halaman Setting Attendance Spot Personnel', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.12 Membuka halaman Setting Attendance Spot Personnel', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.13 Melakukan filter data Personnel', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.14 Mencari data Personnel dengan keyword valid', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.15 Mencari data Personnel dengan keyword tidak valid', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.16 Menambahkan data Personnel secara single', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.16 Menambahkan data Personnel secara single', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.17 Menghapus data Personnel secara single', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.17 Menghapus data Personnel secara single', '2.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.18 Menambahkan data Personnel secara bulk (Add Personnel)', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.18 Menambahkan data Personnel secara bulk (Add Personnel)', '2.png'): (338, 0), # Alert
    ('40. Setting Absent - Attendance Spot', '40.19 Menghapus data Personnel secara bulk (Remove Personnel)', '1.png'): (338, 72),
    ('40. Setting Absent - Attendance Spot', '40.19 Menghapus data Personnel secara bulk (Remove Personnel)', '2.png'): (338, 0), # Alert

    # 41. Setting Absent - Work Pattern
    # Opening
    ('41. Setting Absent - Work Pattern', '41.1 Membuka halaman Work Pattern', '1.png'): (0, 0),
    # Search
    ('41. Setting Absent - Work Pattern', '41.2 Mencari data Work Pattern dengan keyword valid', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.3 Mencari data Work Pattern dengan keyword tidak valid', '1.png'): (338, 72),
    # Add
    ('41. Setting Absent - Work Pattern', '41.4 Menambahkan data Work Pattern', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.4 Menambahkan data Work Pattern', '2.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.4 Menambahkan data Work Pattern', '3.png'): (338, 0), # Alert
    # Add Mandatory
    ('41. Setting Absent - Work Pattern', '41.5 Menambahkan data Work Pattern tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.5 Menambahkan data Work Pattern tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Apply to all
    ('41. Setting Absent - Work Pattern', '41.6 Menerapkan jadwal ke beberapa hari secara sekaligus (Terapkan ke Semua)', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.6 Menerapkan jadwal ke beberapa hari secara sekaligus (Terapkan ke Semua)', '2.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.6 Menerapkan jadwal ke beberapa hari secara sekaligus (Terapkan ke Semua)', '3.png'): (338, 72),
    # Reset
    ('41. Setting Absent - Work Pattern', '41.7 Mereset jadwal pada beberapa hari secara sekaligus (Reset)', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.7 Mereset jadwal pada beberapa hari secara sekaligus (Reset)', '2.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.7 Mereset jadwal pada beberapa hari secara sekaligus (Reset)', '3.png'): (338, 72),
    # Edit
    ('41. Setting Absent - Work Pattern', '41.8 Mengubah data Work Pattern', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.8 Mengubah data Work Pattern', '2.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.8 Mengubah data Work Pattern', '3.png'): (338, 0), # Alert
    # Edit Mandatory
    ('41. Setting Absent - Work Pattern', '41.9 Mengubah data Work Pattern tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.9 Mengubah data Work Pattern tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Detail Modal
    ('41. Setting Absent - Work Pattern', '41.10 Melihat detail Work Pattern', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.10 Melihat detail Work Pattern', '2.png'): (338, 72),
    # Delete
    ('41. Setting Absent - Work Pattern', '41.11 Menghapus data Work Pattern', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.11 Menghapus data Work Pattern', '2.png'): (338, 0), # Alert
    # Setting Personnel
    ('41. Setting Absent - Work Pattern', '41.12 Membuka halaman Setting Work Pattern Personnel', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.12 Membuka halaman Setting Work Pattern Personnel', '2.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.13 Melakukan filter data Personnel', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.14 Mencari data Personnel dengan keyword valid', '1.png'): (338, 72),
    ('41. Setting Absent - Work Pattern', '41.15 Mencari data Personnel dengan keyword tidak valid', '1.png'): (338, 72),
    # 41.16 - 41.19: Collapsed Sidebar (80px)!
    ('41. Setting Absent - Work Pattern', '41.16 Meng-assign data Personnel ke Work Pattern secara single', '1.png'): (80, 72),
    ('41. Setting Absent - Work Pattern', '41.16 Meng-assign data Personnel ke Work Pattern secara single', '2.png'): (80, 72), # Modal confirmation
    ('41. Setting Absent - Work Pattern', '41.16 Meng-assign data Personnel ke Work Pattern secara single', '3.png'): (80, 0), # Alert
    ('41. Setting Absent - Work Pattern', '41.17 Meng-assign data Personnel ke Work Pattern secara bulk', '1.png'): (80, 72),
    ('41. Setting Absent - Work Pattern', '41.17 Meng-assign data Personnel ke Work Pattern secara bulk', '2.png'): (80, 72), # Modal confirmation
    ('41. Setting Absent - Work Pattern', '41.17 Meng-assign data Personnel ke Work Pattern secara bulk', '3.png'): (80, 0), # Alert
    ('41. Setting Absent - Work Pattern', '41.18 Meng-unassign data Personnel dari Work Pattern secara single', '1.png'): (80, 72),
    ('41. Setting Absent - Work Pattern', '41.18 Meng-unassign data Personnel dari Work Pattern secara single', '2.png'): (80, 72), # Modal confirmation
    ('41. Setting Absent - Work Pattern', '41.18 Meng-unassign data Personnel dari Work Pattern secara single', '3.png'): (80, 0), # Alert
    ('41. Setting Absent - Work Pattern', '41.19 Meng-unassign data Personnel dari Work Pattern secara bulk', '1.png'): (80, 72),
    ('41. Setting Absent - Work Pattern', '41.19 Meng-unassign data Personnel dari Work Pattern secara bulk', '2.png'): (80, 72), # Modal confirmation
    ('41. Setting Absent - Work Pattern', '41.19 Meng-unassign data Personnel dari Work Pattern secara bulk', '3.png'): (80, 0), # Alert

    # 42. Setting Absent - Holiday
    # Opening
    ('42. Setting Absent - Holiday', '42.1 Membuka halaman Holiday', '1.png'): (0, 0),
    # Search
    ('42. Setting Absent - Holiday', '42.2 Mencari data Holiday dengan keyword valid', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.3 Mencari data Holiday dengan keyword tidak valid', '1.png'): (338, 72),
    # Add
    ('42. Setting Absent - Holiday', '42.4 Menambahkan data Holiday', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.4 Menambahkan data Holiday', '2.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.4 Menambahkan data Holiday', '3.png'): (338, 0), # Alert
    # Add Mandatory
    ('42. Setting Absent - Holiday', '42.5 Menambahkan data Holiday tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.5 Menambahkan data Holiday tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Edit
    ('42. Setting Absent - Holiday', '42.6 Mengubah data Holiday', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.6 Mengubah data Holiday', '2.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.6 Mengubah data Holiday', '3.png'): (338, 0), # Alert
    # Edit Mandatory
    ('42. Setting Absent - Holiday', '42.7 Mengubah data Holiday tanpa mengisi field mandatory', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.7 Mengubah data Holiday tanpa mengisi field mandatory', '2.png'): (338, 72),
    # Delete
    ('42. Setting Absent - Holiday', '42.8 Menghapus data Holiday', '1.png'): (338, 72),
    ('42. Setting Absent - Holiday', '42.8 Menghapus data Holiday', '2.png'): (338, 0), # Alert
}

print(f"\nTotal crop rules configured: {len(crop_rules)}")

# Execute Cropping from (Original Full) to Active
crop_count = 0
for (mod_name, tc_dir, img_file), (crop_x, crop_y) in crop_rules.items():
    orig_dir = mod_name + ' (Original Full)'
    src_img = os.path.join(base, orig_dir, tc_dir, img_file)
    dst_img = os.path.join(base, mod_name, tc_dir, img_file)
    
    if not os.path.exists(win_p(src_img)):
        print(f"ERROR: Source image not found: {src_img}")
        continue
        
    os.makedirs(win_p(os.path.dirname(dst_img)), exist_ok=True)
    with Image.open(win_p(src_img)) as im:
        w, h = im.size
        if crop_x == 0 and crop_y == 0:
            cropped = im.copy()
            action_desc = "FULL SCREEN"
        else:
            cropped = im.crop((crop_x, crop_y, w, h))
            action_desc = f"CROP (x={crop_x}, y={crop_y})"
            
        cropped.save(win_p(dst_img))
        crop_count += 1
        print(f"[{crop_count:2d}] {mod_name[:14]} | {tc_dir[:25]} / {img_file} -> {action_desc} | {cropped.size}")

print(f"\nSuccessfully cropped {crop_count} images across Modul 40, 41, 42!")
