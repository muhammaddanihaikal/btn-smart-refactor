import os

folder = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile\05. Absent - Clock In\5.1 Melakukan Clock In'
for f in os.listdir(folder):
    p = os.path.join(folder, f)
    with open(p, 'rb') as fp:
        head = fp.read(8)
    print(f"{f}: {head.hex()}")
