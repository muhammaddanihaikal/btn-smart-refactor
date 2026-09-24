import win32com.client
import sys

print("Dispatching Word...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

docx_path = r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx'
print(f"Opening {docx_path}...")
try:
    doc = word.Documents.Open(
        FileName=docx_path,
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        PasswordDocument="",
        PasswordTemplate="",
        Revert=False,
        WritePasswordDocument="",
        WritePasswordTemplate="",
        Format=0,
        Encoding=65001,
        Visible=False,
        OpenAndRepair=False,
        NoEncodingDialog=True
    )
    print("Document successfully opened!")
    pages = doc.ComputeStatistics(2)
    print(f"Total pages: {pages}")
    
    # Let's check paragraphs on each page or export to PDF
    pdf_path = r'd:\Project\BTN\scratch\sit_web.pdf'
    doc.SaveAs2(pdf_path, FileFormat=17)
    print(f"Saved PDF to {pdf_path}")
    doc.Close(False)
except Exception as e:
    print(f"Exception during open: {e}")
finally:
    word.Quit()
    print("Done.")
