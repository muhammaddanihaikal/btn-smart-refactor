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

# 1. Update TOC
doc_p = os.path.join(d_base, r'Hasil Uji\Dokumen_Hasil_Uji_Web.docx')
try:
    import win32com.client
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    wdoc = word.Documents.Open(os.path.abspath(doc_p), ConfirmConversions=False, ReadOnly=False)
    for toc in wdoc.TablesOfContents:
        toc.Update()
    wdoc.Save()
    wdoc.Close(SaveChanges=True)
    word.Quit()
    print("TOC updated successfully!")
except Exception as e:
    print("TOC update skipped/error:", e)

# 2. Sync Document to Drive H and Drive G
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

# 3. Sync Modul 07 active folder
sync_folder = r'Hasil Uji\Screenshot\Web\07. User Authority - Group Role'

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
                if not os.path.exists(d_f) or os.path.getsize(s_f) != os.path.getsize(d_f) or os.path.getmtime(s_f) > os.path.getmtime(d_f):
                    shutil.copy2(s_f, d_f)
                    synced_files += 1
            except Exception as e:
                print(f"[{label}] Copy error for {f}: {e}")
    print(f"[{label}] Synced {synced_files} files in {os.path.basename(src)}")

for dst_root, label in [(h_base, 'Drive H'), (g_base, 'Drive G')]:
    sync_dir(os.path.join(d_base, sync_folder), os.path.join(dst_root, sync_folder), label)

# 4. Hash verification
def get_md5(p):
    h = hashlib.md5()
    with open(win_p(p), 'rb') as f:
        while chunk := f.read(1024 * 1024):
            h.update(chunk)
    return h.hexdigest()

print("\n=== VERIFYING DOCUMENT HASHES ACROSS DRIVES ===")
d_hash = get_md5(src_doc)
print(f"[Local D] Size: {os.path.getsize(src_doc):,} bytes | MD5: {d_hash}")
for dst_root, label in [(h_base, 'Drive H'), (g_base, 'Drive G')]:
    p = os.path.join(dst_root, doc_rel)
    if os.path.exists(win_p(p)):
        print(f"[{label}] Size: {os.path.getsize(p):,} bytes | MD5: {get_md5(p)}")
    else:
        print(f"[{label}] File NOT FOUND!")

print("\nTri-drive sync finished!")
