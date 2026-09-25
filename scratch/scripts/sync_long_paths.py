import os
import shutil

base_h = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'
base_d = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot'
base_g = r'G:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot'

def make_long(p):
    abs_p = os.path.abspath(p)
    if os.name == 'nt' and not abs_p.startswith('\\\\?\\'):
        return '\\\\?\\' + abs_p
    return abs_p

copied_d = 0
copied_g = 0

for root, dirs, files in os.walk(base_h):
    for f in files:
        if f == 'desktop.ini':
            continue
        p_h = os.path.join(root, f)
        rel = os.path.relpath(p_h, base_h)
        
        # Sync to D
        p_d = os.path.join(base_d, rel)
        if not os.path.exists(p_d):
            os.makedirs(os.path.dirname(make_long(p_d)), exist_ok=True)
            shutil.copy2(make_long(p_h), make_long(p_d))
            copied_d += 1
            
        # Sync to G
        p_g = os.path.join(base_g, rel)
        if not os.path.exists(p_g):
            os.makedirs(os.path.dirname(make_long(p_g)), exist_ok=True)
            shutil.copy2(make_long(p_h), make_long(p_g))
            copied_g += 1

print(f"Copied {copied_d} missing files to D")
print(f"Copied {copied_g} missing files to G")
