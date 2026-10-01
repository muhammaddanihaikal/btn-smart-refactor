import openpyxl

excel_path = r'D:\Project\BTN Smart\Refactor\Test Script\Uji Sistem.xlsx'
wb = openpyxl.load_workbook(excel_path)
sheet = wb['TC BTN SMART Mobile']

# Find row for 'Membatalkan prospek melalui menu titik tiga'
target_row = None
for r in range(1, sheet.max_row + 1):
    val = sheet.cell(r, 3).value
    if val and 'Membatalkan prospek' in str(val):
        target_row = r
        break

print(f'Found target row: {target_row}')

if target_row:
    # Insert 3 rows before target_row
    sheet.insert_rows(target_row, 3)
    
    # 1. Row target_row (10.21)
    sheet.cell(target_row, 1).value = 'Prospek & Nasabah'
    sheet.cell(target_row, 2).value = 'Aktifitas Prospek'
    sheet.cell(target_row, 3).value = 'Membuka tab Top Up data pada halaman aktifitas Marketing prospek ETB'
    sheet.cell(target_row, 4).value = 'Positive Case'
    sheet.cell(target_row, 5).value = 'Normal'
    sheet.cell(target_row, 6).value = '1. Klik tab Top Up dana pada halaman aktifitas marketing Prospek ETB.'
    sheet.cell(target_row, 7).value = '1. Berhasil berpindah ke tab Top Up dana dan menampilkan form Top Up dana.'
    
    # 2. Row target_row + 1 (10.22)
    sheet.cell(target_row + 1, 1).value = 'Prospek & Nasabah'
    sheet.cell(target_row + 1, 2).value = 'Aktifitas Prospek'
    sheet.cell(target_row + 1, 3).value = 'Melakukan Submit Top Up dana dengan data valid'
    sheet.cell(target_row + 1, 4).value = 'Positive Case'
    sheet.cell(target_row + 1, 5).value = 'High'
    sheet.cell(target_row + 1, 6).value = '1. Mengisi seluruh field pada form Top Up dana dengan data valid.\n2. Klik button Submit.'
    sheet.cell(target_row + 1, 7).value = '1. Berhasil mengisi seluruh field form Top Up dana.\n2. Berhasil men-submit form dan menyimpan data Open Account.'
    
    # 3. Row target_row + 2 (10.23)
    sheet.cell(target_row + 2, 1).value = 'Prospek & Nasabah'
    sheet.cell(target_row + 2, 2).value = 'Aktifitas Prospek'
    sheet.cell(target_row + 2, 3).value = 'Melakukan Submit Top Up dana tanpa field mandatory'
    sheet.cell(target_row + 2, 4).value = 'Negative Case'
    sheet.cell(target_row + 2, 5).value = 'Normal'
    sheet.cell(target_row + 2, 6).value = '1. Mengosongkan seluruh field mandatory pada form Top Up dana.\n2. Klik button Submit.'
    sheet.cell(target_row + 2, 7).value = '1. Berhasil mengosongkan field mandatory.\n2. Menampilkan pesan validasi bahwa field mandatory wajib diisi.'
    
    wb.save(excel_path)
    print('Successfully updated Test Script/Uji Sistem.xlsx!')
else:
    print('Target row not found!')
