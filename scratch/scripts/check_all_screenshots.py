import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'
for root, dirs, files in os.walk(base_dir):
    imgs = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    if imgs:
        print(f"{os.path.basename(root)}: {imgs}")
