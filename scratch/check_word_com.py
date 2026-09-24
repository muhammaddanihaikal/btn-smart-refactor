import win32com.client
import os, sys

docx_path = os.path.abspath(r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx')
print("Initializing Word...")
sys.stdout.flush()

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

try:
    print(f"Opening {docx_path}...")
    sys.stdout.flush()
    doc = word.Documents.Open(docx_path, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
    print("Opened successfully!")
    sys.stdout.flush()
    pages = doc.ComputeStatistics(2)
    print(f"Total pages: {pages}")
    doc.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    word.Quit()
    print("Word quit.")
