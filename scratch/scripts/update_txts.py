import os
import re

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\01. Login'

updates = {
    '1.2 Melakukan login dengan email dan password valid': {
        'steps': "1. Pilih Type.\n2. Isi field Email dengan email terdaftar.\n3. Isi field Password dengan password yang valid.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengisi Email.\n3. Berhasil mengisi Password.\n4. Berhasil melakukan login dan diarahkan ke halaman dashboard."
    },
    '1.3 Melakukan login dengan email tidak terdaftar': {
        'steps': "1. Pilih Type.\n2. Isi field Email dengan email yang tidak terdaftar.\n3. Isi field Password dengan password bebas.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengisi Email.\n3. Berhasil mengisi Password.\n4. Menampilkan pesan error email tidak terdaftar."
    },
    '1.4 Melakukan login dengan password salah': {
        'steps': "1. Pilih Type.\n2. Isi field Email dengan email terdaftar.\n3. Isi field Password dengan password yang salah.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengisi Email.\n3. Berhasil mengisi Password.\n4. Menampilkan pesan error password salah."
    },
    '1.5 Melakukan login tanpa mengisi email': {
        'steps': "1. Pilih Type.\n2. Kosongkan field Email.\n3. Isi field Password dengan password yang valid.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengosongkan Email.\n3. Berhasil mengisi Password.\n4. Menampilkan pesan error field mandatory."
    },
    '1.6 Melakukan login tanpa mengisi password': {
        'steps': "1. Pilih Type.\n2. Isi field Email dengan email terdaftar.\n3. Kosongkan field Password.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengisi Email.\n3. Berhasil mengosongkan Password.\n4. Menampilkan pesan error field mandatory."
    },
    '1.7 Melakukan login tanpa mengisi email dan password': {
        'steps': "1. Pilih Type.\n2. Kosongkan field Email.\n3. Kosongkan field Password.\n4. Klik button LOGIN.",
        'expected': "1. Berhasil memilih Type.\n2. Berhasil mengosongkan Email.\n3. Berhasil mengosongkan Password.\n4. Menampilkan pesan error field mandatory."
    }
}

for folder_name, data in updates.items():
    folder_path = os.path.join(base_dir, folder_name)
    txt_path = os.path.join(folder_path, folder_name + '.txt')
    
    if os.path.exists(txt_path):
        with open(txt_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        content = re.sub(
            r'(STEPS \(LANGKAH-LANGKAH PENGUJIAN\):\n={50}\n).*?(?=\n={50}\nEXPECTED RESULTS)',
            r'\g<1>' + data["steps"],
            content,
            flags=re.DOTALL
        )
        
        content = re.sub(
            r'(EXPECTED RESULTS \(HASIL YANG DIHARAPKAN\):\n={50}\n).*',
            r'\g<1>' + data["expected"] + '\n',
            content,
            flags=re.DOTALL
        )
        
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Updated {txt_path}")
