import docx
import copy
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc_path = r'D:/Project/BTN Smart/Refactor/Hasil Uji/Dokumen_Hasil_Uji_Mobile.docx'
doc = docx.Document(doc_path)

# Table 133 currently is 10.21 Membatalkan prospek
t133 = doc.tables[133]

# We will create 3 new tables by deepcopying t133._tbl
tbl_parent = t133._tbl.getparent()

# 1. Update t133 to 10.21
t133.rows[0].cells[0].text = '10.21'
t133.rows[0].cells[1].text = ' Membuka tab Top Up data pada halaman aktifitas Marketing prospek ETB'
t133.rows[1].cells[0].text = ''
t133.rows[1].cells[1].text = ''
t133.rows[2].cells[0].text = ''
t133.rows[2].cells[1].text = 'Hasil yang diharapkan [STATUS]:\nBerhasil berpindah ke tab Top Up dana dan menampilkan form Top Up dana.'

# Apply Arial font
for r in t133.rows:
    for c in r.cells:
        for p in c.paragraphs:
            for run in p.runs:
                run.font.name = 'Arial'

# Helper to configure a cloned table
def configure_table(tbl_element, tc_id, tc_title, expected_text):
    tbl = docx.table.Table(tbl_element, doc)
    tbl.rows[0].cells[0].text = tc_id
    tbl.rows[0].cells[1].text = ' ' + tc_title
    tbl.rows[1].cells[0].text = ''
    tbl.rows[1].cells[1].text = ''
    tbl.rows[2].cells[0].text = ''
    tbl.rows[2].cells[1].text = 'Hasil yang diharapkan [STATUS]:\n' + expected_text
    for r in tbl.rows:
        for c in r.cells:
            for p in c.paragraphs:
                for run in p.runs:
                    run.font.name = 'Arial'

# 2. Table 10.22
tbl_22_elem = copy.deepcopy(t133._tbl)
configure_table(tbl_22_elem, '10.22', 'Melakukan Submit Top Up dana dengan data valid', 'Berhasil men-submit form dan menyimpan data Open Account.')
t133._tbl.addnext(tbl_22_elem)

# 3. Table 10.23
tbl_23_elem = copy.deepcopy(t133._tbl)
configure_table(tbl_23_elem, '10.23', 'Melakukan Submit Top Up dana tanpa field mandatory', 'Menampilkan pesan validasi bahwa field mandatory wajib diisi.')
tbl_22_elem.addnext(tbl_23_elem)

# 4. Table 10.24
tbl_24_elem = copy.deepcopy(t133._tbl)
configure_table(tbl_24_elem, '10.24', 'Membatalkan prospek melalui menu titik tiga', 'Berhasil membatalkan prospek dan prospek tidak dapat dilanjutkan kembali.')
tbl_23_elem.addnext(tbl_24_elem)

doc.save(doc_path)
print('Successfully added 10.21-10.24 to Dokumen_Hasil_Uji_Mobile.docx!')
