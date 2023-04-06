import xlwings as xw
# choose if excel is visible or not
app = xw.App(visible=False)
# open workbook and worksheet
wb = xw.Book('data.xlsx')
sheet = wb.sheets['Tabelle1']
# set cell B6  to "My new Data"
sheet['B6'].value = "My new Data"
# close workbook and quit app
wb.close()
app.quit()