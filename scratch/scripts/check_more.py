import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

for root, dirs, files in os.walk(base_dir):
    rel = os.path.relpath(root, base_dir)
    if rel.startswith('01. Login') or rel.startswith('02. Profile') or rel.startswith('03. Sales Tracking Activity'):
        continue
    for f in files:
        if f.lower().endswith(('.jpg', '.jpeg', '.png')) or ('.' not in f and f != 'desktop.ini'):
            print(f"Found in {rel}: {f}")
