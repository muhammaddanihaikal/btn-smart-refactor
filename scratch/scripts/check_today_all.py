import os
import datetime

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

today_all = []
today_str = "2026-09-25"

for platform in ['Mobile', 'Web']:
    plat_dir = os.path.join(base_dir, platform)
    if not os.path.exists(plat_dir):
        continue
    for root, dirs, files in os.walk(plat_dir):
        for f in files:
            if f.endswith('.txt') or f == 'desktop.ini':
                continue
            p = os.path.join(root, f)
            mtime = os.path.getmtime(p)
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            if dt_str.startswith(today_str):
                rel = os.path.relpath(p, base_dir)
                size = os.path.getsize(p)
                today_all.append((rel, size, dt_str))

print(f"Total files uploaded anytime today ({today_str}): {len(today_all)}")
for rel, size, dt_str in sorted(today_all, key=lambda x: x[2]):
    print(f" - {rel} ({size:,} B) - {dt_str}")
