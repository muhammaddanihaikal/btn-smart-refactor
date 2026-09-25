import os

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

all_images = {}

for platform in ['Mobile', 'Web']:
    plat_dir = os.path.join(base_dir, platform)
    if not os.path.exists(plat_dir):
        continue
    for root, dirs, files in os.walk(plat_dir):
        rel = os.path.relpath(root, base_dir)
        imgs = [f for f in files if not f.endswith('.txt') and f != 'desktop.ini']
        if imgs:
            all_images[rel] = imgs

print(f"Total folders with screenshots across Mobile & Web: {len(all_images)}")
for folder in sorted(all_images.keys()):
    print(f"[{folder}]: {all_images[folder]}")
