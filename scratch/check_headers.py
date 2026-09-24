import zipfile

with zipfile.ZipFile(r'd:\Project\BTN\SIT\Form Script - Skenario SIT BTN SMART Upgrade Server.docx') as z:
    for name in z.namelist():
        if 'header' in name or 'footer' in name:
            print(name, len(z.read(name)))
