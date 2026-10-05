import os
import shutil
import time
import hashlib

def win_p(p):
    ap = os.path.abspath(p)
    return '\\\\?\\' + ap if not ap.startswith('\\\\?\\') else p

d_base = r'D:\Project\BTN Smart\Refactor'
h_base = r'H:\My Drive\Zegen\BTN Smart\Refactor'
g_base = r'G:\My Drive\Zegen\BTN Smart\Refactor'

# 1. Sync Document
doc_rel = r'Hasil Uji\Dokumen_Hasil_Uji_Web.docx'
src_doc = os.path.join(d_base, doc_rel)

for dst_root, label in [(h_base, 'Drive H'), (g_base, 'Drive G')]:
    dst_doc = os.path.join(dst_root, doc_rel)
    success = False
    for attempt in range(5):
        try:
            shutil.copy2(win_p(src_doc), win_p(dst_doc))
            print(f"[{label}] Successfully synced {doc_rel} ({os.path.getsize(src_doc):,} bytes)")
            success = True
            break
        except Exception as e:
            print(f"[{label}] Attempt {attempt+1} failed: {e}. Retrying in 2s...")
            time.sleep(2)
    if not success:
        print(f"[{label}] ERROR: Failed to sync {doc_rel}")

# 2. Sync Modul 06 Folders
sync_folders = [
    r'Hasil Uji\Screenshot\Web\06. User Authority - User',
    r'Hasil Uji\Screenshot\Web\06. User Authority - User (Original Full)',
]

def sync_dir(src, dst, label):
    os.makedirs(win_p(dst), exist_ok=True)
    synced_files = 0
    for root, dirs, files in os.walk(win_p(src)):
        rel = os.path.relpath(root, win_p(src))
        target_dir = os.path.join(win_p(dst), rel) if rel != '.' else win_p(dst)
        os.makedirs(target_dir, exist_ok=True)
        for f in files:
            if f.lower() == 'desktop.ini':
                continue
            s_f = os.path.join(root, f)
            d_f = os.path.join(target_dir, f)
            try:
                if not os.path.exists(d_f) or os.path.getsize(s_f) != os.path.getsize(d_f):
                    shutil.copy2(s_f, d_f)
                    synced_files += 1
            except Exception as e:
                print(f"[{label}] Error copying {f}: {e}")
    if synced_files > 0:
        print(f"[{label}] Synced {synced_files} files in {os.path.basename(src)}")

for fld in sync_folders:
    src_fld = os.path.join(d_base, fld)
    if os.path.exists(win_p(src_fld)):
        for dst_root, label in [(h_base, 'Drive H'), (g_base, 'Drive G')]:
            dst_fld = os.path.join(dst_root, fld)
            sync_dir(src_fld, dst_fld, label)

# 3. Verify Hashes of Document
print("\n=== VERIFYING DOCUMENT HASHES ACROSS DRIVES ===")
for root_path, label in [(d_base, 'Local D'), (h_base, 'Drive H'), (g_base, 'Drive G')]:
    fp = os.path.join(root_path, doc_rel)
    if os.path.exists(win_p(fp)):
        with open(win_p(fp), 'rb') as f:
            h = hashlib.md5(f.read()).hexdigest()
        sz = os.path.getsize(win_p(fp))
        print(f"[{label:7s}] Size: {sz:,} bytes | MD5: {h}")
    else:
        print(f"[{label:7s}] MISSING!")

print("\nTri-drive sync finished!")
