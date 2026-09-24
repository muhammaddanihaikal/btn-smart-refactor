import os
import datetime

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

login_tcs = {}
other_tcs = {}

for root, dirs, files in os.walk(base_dir):
    rel = os.path.relpath(root, base_dir)
    imgs = []
    for f in files:
        if f.lower().endswith(('.jpg', '.jpeg', '.png')) or ('.' not in f and f != 'desktop.ini'):
            p = os.path.join(root, f)
            mtime = os.path.getmtime(p)
            size = os.path.getsize(p)
            dt = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            imgs.append((f, size, dt))
            
    if imgs:
        if rel.startswith('01. Login'):
            login_tcs[rel] = imgs
        else:
            other_tcs[rel] = imgs

print("=== 01. LOGIN SCREENSHOTS ===")
for folder in sorted(login_tcs.keys()):
    print(f"\n[Folder] {folder}")
    for name, size, dt in login_tcs[folder]:
        print(f"   - {name} ({size:,} B, {dt})")

print("\n=== OTHER MODULES SCREENSHOTS ===")
for folder in sorted(other_tcs.keys()):
    print(f"\n[Folder] {folder}")
    for name, size, dt in other_tcs[folder]:
        print(f"   - {name} ({size:,} B, {dt})")
