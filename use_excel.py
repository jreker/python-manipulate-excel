import xlwings as xw

# choose if excel is visible or not
app = xw.App(visible=False)

wb = xw.Book('data.xlsx')
sheet = wb.sheets['Tabelle1']

# change data to lower the avg
sheet['L2'].value = "7"
sheet['Q2'].value = "5"

# print result of formula
print(sheet['B4'].value)

# change input data 
sheet['L2'].value = "35"
sheet['Q2'].value = "31"

# print result of formula after change
print(sheet['B4'].value)

wb.close()
app.quit()