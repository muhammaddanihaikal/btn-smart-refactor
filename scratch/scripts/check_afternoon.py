import os
import datetime

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

cutoff = datetime.datetime(2026, 9, 25, 10, 36, 0)
recent_files = []

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
            dt = datetime.datetime.fromtimestamp(mtime)
            if dt > cutoff:
                rel = os.path.relpath(p, base_dir)
                size = os.path.getsize(p)
                dt_str = dt.strftime('%Y-%m-%d %H:%M:%S')
                recent_files.append((rel, size, dt_str))

print(f"Total files uploaded/modified after 10:36 today: {len(recent_files)}")
for rel, size, dt_str in sorted(recent_files, key=lambda x: x[2]):
    print(f" - {rel} ({size:,} B) - {dt_str}")
