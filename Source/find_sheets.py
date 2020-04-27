import openpyxl as openpyxl

#fp = 'C:/Users/rjw3/OneDrive/Desktop/Test.xlsx'
def find_sheets(fp):
    
    wb = openpyxl.load_workbook(fp)
    
    sheets = wb.sheetnames
    wb.close()
    return sheets