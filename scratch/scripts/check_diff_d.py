import os

base_h = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'
base_d = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot'

missing_in_d = []
for root, dirs, files in os.walk(base_h):
    for f in files:
        if f == 'desktop.ini':
            continue
        p_h = os.path.join(root, f)
        rel = os.path.relpath(p_h, base_h)
        p_d = os.path.join(base_d, rel)
        if not os.path.exists(p_d):
            missing_in_d.append(rel)

print(f"Total files in H missing in D: {len(missing_in_d)}")
for m in missing_in_d[:10]:
    print(f" - {m}")
