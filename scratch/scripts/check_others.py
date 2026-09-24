import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

for root, dirs, files in os.walk(base_dir):
    for f in files:
        if not f.endswith('.txt'):
            rel = os.path.relpath(os.path.join(root, f), base_dir)
            if not rel.startswith('01. Login') and not rel.startswith('02. Profile') and not rel.startswith('03. Sales Tracking Activity'):
                print(rel)
