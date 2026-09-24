from PIL import Image
import os

base = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'
for tc in [
    r'02. Profile\2.1 Membuka halaman profile',
    r'02. Profile\2.2 Membuka halaman Informasi Pribadi',
    r'02. Profile\2.3 Mengubah foto profil',
    r'03. Sales Tracking Activity - Visit In\3.1 Melakukan Visit in menggunakan sales Funding',
    r'03. Sales Tracking Activity - Visit In\3.2 Melakukan Visit in menggunakan sales Funding'
]:
    p = os.path.join(base, tc)
    for f in sorted(os.listdir(p)):
        if f.lower().endswith(('.jpeg', '.jpg', '.png')):
            with Image.open(os.path.join(p, f)) as img:
                print(f"{tc} -> {f}: {img.size}")
