import win32com.client

docx_path = r'd:\Project\BTN\SIT\Document SIT BTN Smart\SIT BTN SMART Web.docx'
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0

try:
    doc = word.Documents.Open(docx_path)
    total_pages = doc.ComputeStatistics(2)
    print(f"Total pages: {total_pages}")
    
    # Check each page content
    # In Word, wdGoToPage = 1, wdGoToAbsolute = 1
    # We can inspect headings and pages
    headings = []
    for p in doc.Paragraphs:
        txt = p.Range.Text.strip()
        if txt and ('Modul ' in txt or 'Module ' in txt) and not txt.startswith('Modul :'):
            # get page number
            pg = p.Range.Information(3) # 3 = wdActiveEndPageNumber
            headings.append((pg, txt))
            
    print(f"Found {len(headings)} section headings across document.")
    for pg, h in headings[:15]:
        print(f"  Page {pg:3d}: {h}")
        
    # Check if there are consecutive page jumps without headings or large tables
    for i in range(len(headings) - 1):
        pg1, h1 = headings[i]
        pg2, h2 = headings[i+1]
        diff = pg2 - pg1
        # If difference is more than 1, let's see how many pages that section takes
        # e.g. section with 37 TCs will take 3 or 4 pages.
        # But if a section with 3 TCs takes 3 pages, that might be a blank page!
    
    doc.Close(False)
finally:
    word.Quit()
print("Done!")
