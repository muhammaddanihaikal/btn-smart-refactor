import win32com.client
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
tpl_path = r'd:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx'
print(f"Opening template {tpl_path}...")
try:
    doc = word.Documents.Open(tpl_path)
    print("Template opened successfully!")
    print("Pages:", doc.ComputeStatistics(2))
    doc.Close(False)
finally:
    word.Quit()
print("Done!")
