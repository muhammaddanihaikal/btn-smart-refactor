import win32com.client
import os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False

doc_path = r'D:\Project\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Mobile.docx'
pdf_path = r'D:\Project\BTN Smart\Refactor\scratch\preview_1_1.pdf'

doc = word.Documents.Open(doc_path)
doc.SaveAs(pdf_path, FileFormat=17)
doc.Close()
word.Quit()
print('PDF created successfully')
