import zipfile, os, shutil
import win32com.client

# Let's test opening a copy of SIT BTN SMART Web.docx with original header1.xml from template
test_docx = r'd:\Project\BTN\scratch\test_no_postprocess.docx'
shutil.copy2(r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx', test_docx)

# Replace header1.xml and settings.xml with the ones from template
with zipfile.ZipFile(r'd:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx', 'r') as z_tpl:
    h1 = z_tpl.read('word/header1.xml')
    s = z_tpl.read('word/settings.xml')

tmp = test_docx + '.tmp'
with zipfile.ZipFile(test_docx, 'r') as z_in, zipfile.ZipFile(tmp, 'w') as z_out:
    for item in z_in.infolist():
        if item.filename == 'word/header1.xml':
            z_out.writestr(item, h1)
        elif item.filename == 'word/settings.xml':
            z_out.writestr(item, s)
        else:
            z_out.writestr(item, z_in.read(item.filename))
shutil.move(tmp, test_docx)

print("Opening test docx with template header/settings in Word...")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(test_docx)
    print("Opened successfully! Total pages:", doc.ComputeStatistics(2))
    doc.Close(False)
except Exception as e:
    print("Error:", e)
finally:
    word.Quit()
print("Done!")
