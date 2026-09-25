import os

SCREENSHOT_BASE = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile'

tcs = ['5.1', '5.2', '5.3', '5.4', '5.5', '5.6', '5.7', '6.1', '6.2', '6.3']

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
