import os

local_dir = r'd:\Project\BTN\Hasil Uji\Screenshot\Web'
drive_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Web'

for name, base_dir in [('Local', local_dir), ('Drive H', drive_dir)]:
    if not os.path.exists(base_dir):
        print(f"{name}: directory does not exist: {base_dir}")
        continue
    items = sorted(os.listdir(base_dir))
    print(f"{name}: total top-level items: {len(items)}")
    print(f"  First 5: {items[:5]}")
    print(f"  Last 5:  {items[-5:]}")
    
    # Check if there are image files anywhere
    img_files = []
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            if f.lower().endswith(('.png', '.jpg', '.jpeg')):
                img_files.append(os.path.join(root, f))
    print(f"  Total images found: {len(img_files)}")
    for img in img_files[:5]:
        print(f"    {img}")
