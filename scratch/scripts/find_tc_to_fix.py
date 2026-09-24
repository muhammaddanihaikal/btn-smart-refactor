import os
import glob

base_dir = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

matches = []

for root, dirs, files in os.walk(base_dir):
    for file in files:
        if file.endswith('.txt'):
            file_path = os.path.join(root, file)
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                has_buka = "Buka aplikasi BTN Smart" in content
                has_non_employee = "Non Employee" in content
                
                if has_buka or has_non_employee:
                    # extract TC number and name from the folder name
                    tc_name = os.path.basename(root)
                    matches.append({
                        'tc': tc_name,
                        'has_buka': has_buka,
                        'has_non_employee': has_non_employee,
                        'path': file_path
                    })
            except Exception as e:
                pass

matches.sort(key=lambda x: x['tc'])
for m in matches:
    print(f"TC: {m['tc']}")
    if m['has_buka']:
        print("  - Mengandung 'Buka aplikasi BTN Smart'")
    if m['has_non_employee']:
        print("  - Mengandung 'Non Employee'")
