import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

results = []

for root, dirs, files in os.walk(base_dir):
    rel_path = os.path.relpath(root, base_dir)
    if rel_path.startswith('01. Login'):
        continue
        
    imgs = []
    for f in files:
        if f.lower().endswith(('.jpg', '.jpeg', '.png')) or ('.' not in f and f != 'desktop.ini'):
            p = os.path.join(root, f)
            mtime = os.path.getmtime(p)
            size = os.path.getsize(p)
            import datetime
            dt = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            imgs.append((f, size, dt))
            
    if imgs:
        results.append((rel_path, imgs))

print(f"Total non-login folders with images: {len(results)}")
for folder, img_list in results:
    print(f"\n[Folder] {folder}:")
    for name, size, dt in img_list:
        print(f"   - {name} ({size:,} bytes, modified: {dt})")
