import zipfile, io, shutil
import lxml.etree as etree

TEMPLATE_PATH = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Template Dokumen Hasil Uji.docx'
TARGET_PATH   = r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'

print("=== FIXING COVER PAGE LOGOS ===")

# 1. Read template's P0, document.xml.rels, and the media files for image1.png and image2.jpg
with zipfile.ZipFile(TEMPLATE_PATH, 'r') as z_tpl:
    tpl_doc_xml = etree.fromstring(z_tpl.read('word/document.xml'))
    tpl_rels_xml = etree.fromstring(z_tpl.read('word/_rels/document.xml.rels'))
    
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    tpl_p0 = tpl_doc_xml.find(f'.//{W}body/{W}p')
    
    img1_bytes = z_tpl.read('word/media/image1.png')
    img2_bytes = z_tpl.read('word/media/image2.jpg')

# In template, tpl_p0 uses rId8 -> media/image1.png and rId9 -> media/image2.jpg
# Let's inspect target zip
with zipfile.ZipFile(TARGET_PATH, 'r') as z_tgt:
    tgt_doc_xml = etree.fromstring(z_tgt.read('word/document.xml'))
    tgt_rels_xml = etree.fromstring(z_tgt.read('word/_rels/document.xml.rels'))
    
    # Check all existing parts
    all_files = {item.filename: z_tgt.read(item.filename) for item in z_tgt.infolist()}

# Replace target P0 with template P0
tgt_p0 = tgt_doc_xml.find(f'.//{W}body/{W}p')
tgt_body = tgt_doc_xml.find(f'.//{W}body')
tgt_p0_index = tgt_body.index(tgt_p0)
tgt_body.remove(tgt_p0)
tgt_body.insert(tgt_p0_index, tpl_p0)

# Ensure rId8 and rId9 exist in tgt_rels_xml and point to media/image1.png and media/image2.jpg
R_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

# Check if rId8 exists, if so update or add
def set_rel(rels_root, r_id, target, target_type):
    for rel in rels_root.findall(f'.//{{{R_NS}}}Relationship'):
        if rel.attrib.get('Id') == r_id:
            rel.attrib['Target'] = target
            rel.attrib['Type'] = target_type
            return
    # Not found, add new
    new_rel = etree.SubElement(rels_root, f'{{{R_NS}}}Relationship')
    new_rel.attrib['Id'] = r_id
    new_rel.attrib['Type'] = target_type
    new_rel.attrib['Target'] = target

set_rel(tgt_rels_xml, 'rId8', 'media/image1.png', IMAGE_TYPE)
set_rel(tgt_rels_xml, 'rId9', 'media/image2.jpg', IMAGE_TYPE)

# Update all_files
all_files['word/document.xml'] = etree.tostring(tgt_doc_xml, xml_declaration=True, encoding='UTF-8', standalone='yes')
all_files['word/_rels/document.xml.rels'] = etree.tostring(tgt_rels_xml, xml_declaration=True, encoding='UTF-8', standalone='yes')
all_files['word/media/image1.png'] = img1_bytes
all_files['word/media/image2.jpg'] = img2_bytes

# Write back to TARGET_PATH
backup_target = TARGET_PATH.replace('.docx', '_BEFORE_COVER_FIX.docx')
shutil.copy2(TARGET_PATH, backup_target)
print("Backup created at:", backup_target)

with zipfile.ZipFile(TARGET_PATH, 'w', zipfile.ZIP_DEFLATED) as z_out:
    for fname, data in all_files.items():
        z_out.writestr(fname, data)

print("TARGET_PATH updated successfully with restored cover logos!")
