import os

folder_1_4 = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\01. Login\1.4 Melakukan login dengan password salah'
for f in ['1', '2']:
    p = os.path.join(folder_1_4, f)
    if os.path.exists(p):
        with open(p, 'rb') as fp:
            header = fp.read(16)
            print(f"{f}: header = {header.hex()}")
