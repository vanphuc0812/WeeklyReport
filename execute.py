from xlrd import open_workbook


wb = open_workbook('Report.xls')
aes_sheet = wb.sheet_by_name('AES_Config')
print(aes_sheet.ncols)
print(aes_sheet.nrows)