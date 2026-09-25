from PIL import Image
import os

base = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\05. Absent - Clock In'
for tc in ['5.1 Melakukan Clock In', '5.2 Melakukan Clock In']:
    p = os.path.join(base, tc)
    for f in sorted(os.listdir(p)):
        if f.lower().endswith(('.jpg', '.jpeg', '.png')):
            with Image.open(os.path.join(p, f)) as img:
                print(f"{tc} -> {f}: {img.size}")
