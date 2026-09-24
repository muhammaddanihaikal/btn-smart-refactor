import os
import shutil

base_login = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\01. Login'

for root, dirs, files in os.walk(base_login):
    for f in files:
        old_path = os.path.join(root, f)
        
        # 1. Handle missing extensions
        if not '.' in f and os.path.isfile(old_path):
            with open(old_path, 'rb') as fp:
                head = fp.read(4)
            if head.startswith(b'\xff\xd8'):
                new_name = f + '.jpg'
                new_path = os.path.join(root, new_name)
                os.rename(old_path, new_path)
                print(f"Added .jpg extension: {f} -> {new_name}")
                f = new_name
                old_path = new_path
            elif head.startswith(b'\x89PNG'):
                new_name = f + '.png'
                new_path = os.path.join(root, new_name)
                os.rename(old_path, new_path)
                print(f"Added .png extension: {f} -> {new_name}")
                f = new_name
                old_path = new_path
                
        # 2. Handle ' (1)' suffix from browser downloads
        if ' (1)' in f:
            clean_name = f.replace(' (1)', '')
            clean_path = os.path.join(root, clean_name)
            if not os.path.exists(clean_path):
                os.rename(old_path, clean_path)
                print(f"Renamed browser download: {f} -> {clean_name}")
            else:
                # If target exists, overwrite with newer one
                os.remove(clean_path)
                os.rename(old_path, clean_path)
                print(f"Replaced {clean_name} with {f}")
