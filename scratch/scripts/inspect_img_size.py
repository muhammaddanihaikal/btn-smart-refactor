from PIL import Image
import os

folder = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\01. Login\1.2 Melakukan login dengan email dan password valid'
for f in ['1.jpg', '2.jpg']:
    p = os.path.join(folder, f)
    with Image.open(p) as img:
        print(f"{f}: size={img.size}, mode={img.mode}")
