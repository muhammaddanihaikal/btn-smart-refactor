import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

for root, dirs, files in os.walk(base_dir):
    rel = os.path.relpath(root, base_dir)
    imgs = [f for f in files if f.endswith(('.jpg', '.jpeg', '.png'))]
    if imgs and ('06. Absent' in rel or '05. Absent' in rel):
        print(f"[{rel}]: {imgs}")
