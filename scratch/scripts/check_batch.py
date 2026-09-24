import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

for sub in ['01. Login', '02. Profile']:
    sub_dir = os.path.join(base_dir, sub)
    print(f"=== {sub} ===")
    for root, dirs, files in os.walk(sub_dir):
        rel = os.path.relpath(root, base_dir)
        imgs = [f for f in files if not f.endswith('.txt') and f != 'desktop.ini']
        if imgs:
            print(f"[{rel}]: {imgs}")
