import os

SCREENSHOT_BASE = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

tcs = [
    '1.12', '1.13', '1.14', '1.15',
    '2.3', '2.4', '2.5', '2.6', '2.7', '2.8', '2.9', '2.10', '2.11', '2.12', '2.13'
]

for tc_num in tcs:
    tc_folder = None
    for root, dirs, files in os.walk(SCREENSHOT_BASE):
        for d in dirs:
            if d.startswith(tc_num + " ") or d == tc_num:
                tc_folder = os.path.join(root, d)
                break
        if tc_folder:
            break
    print(f"TC {tc_num} -> {tc_folder}")
