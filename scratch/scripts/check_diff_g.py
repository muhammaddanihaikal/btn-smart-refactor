import os

base_h = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'
base_g = r'G:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

missing_in_g = []
for root, dirs, files in os.walk(base_h):
    for f in files:
        if f == 'desktop.ini':
            continue
        p_h = os.path.join(root, f)
        rel = os.path.relpath(p_h, base_h)
        p_g = os.path.join(base_g, rel)
        if not os.path.exists(p_g):
            missing_in_g.append(rel)

print(f"Total files in H missing in G: {len(missing_in_g)}")
for m in missing_in_g[:10]:
    print(f" - {m}")
