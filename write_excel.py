import xlwings as xw

# choose if excel is visible or not
app = xw.App(visible=False)

wb = xw.Book('data.xlsx')
sheet = wb.sheets['Tabelle1']

# set cell b6  to "My new Data"
sheet['B6'].value = "My new Data"

wb.close()
app.quit()