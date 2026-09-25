import os
import datetime

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

today_files = []
recent_files = []
cutoff = datetime.datetime(2026, 9, 24, 17, 30, 0)
today_date = "2026-09-25"

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
            dt_str = dt.strftime('%Y-%m-%d %H:%M:%S')
            rel = os.path.relpath(p, base_dir)
            size = os.path.getsize(p)
            
            if dt_str.startswith(today_date):
                today_files.append((rel, size, dt_str))
            elif dt > cutoff:
                recent_files.append((rel, size, dt_str))

print(f"=== FILES UPLOADED TODAY ({today_date}) ===")
if today_files:
    for rel, size, dt_str in today_files:
        print(f"[Today] {rel} ({size:,} B) - {dt_str}")
else:
    print("None found today so far.")

print(f"\n=== FILES UPLOADED AFTER YESTERDAY 17:30 ===")
if recent_files:
    for rel, size, dt_str in recent_files:
        print(f"[Recent] {rel} ({size:,} B) - {dt_str}")
else:
    print("None found after yesterday 17:30.")
