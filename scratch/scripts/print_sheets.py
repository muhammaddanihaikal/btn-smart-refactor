import openpyxl
wb = openpyxl.load_workbook(r'H:\My Drive\Zegen\BTN Smart\Refactor\Test Script\Test Script BTN Smart Refactor.xlsx', read_only=True)
print('Sheet names:', wb.sheetnames)
