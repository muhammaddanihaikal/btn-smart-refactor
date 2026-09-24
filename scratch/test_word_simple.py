import win32com.client
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
doc = word.Documents.Add()
doc.Content.Text = "Hello World"
doc.SaveAs(r'd:\Project\BTN\scratch\hello.docx')
doc.SaveAs2(r'd:\Project\BTN\scratch\hello.pdf', FileFormat=17)
doc.Close(False)
word.Quit()
print("Success!")
