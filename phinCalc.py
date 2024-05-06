#import libraries 
from openpyxl import load_workbook



excelFile = 'phinCalc.xlsx'



wb = load_workbook(excelFile, data_only = True)
sheet = wb.worksheets[0]


information = sheet['E7':'E10']
income = sheet['D39':'E139']
expense = sheet['F39':'G139']


interestRateOnBalance = information[2][0].value/100.0 + 1
balanceAmount = information[3][0].value


taxRate = 1.0 - 30.0/100.0 #todo: hook up variable tax rate calculation


balance = []

if len(income) != len(expense):
	print('ERROR: income length does not match expense')
	quit()

for index in range(len(income)):

	balanceAmount = balanceAmount * interestRateOnBalance


	yearlyIncome = 0.0

	incomeAmount = income[index][0].value
	if incomeAmount != None:
		yearlyIncome = yearlyIncome + incomeAmount

	incomeAmount = income[index][1].value
	if incomeAmount != None:
		incomeAmount = 12 * incomeAmount
		yearlyIncome = yearlyIncome + incomeAmount

	yearlyIncome = taxRate * yearlyIncome


	yearlyExpense = 0.0

	expenseAmount = expense[index][0].value
	if expenseAmount != None:
		yearlyExpense = yearlyExpense + expenseAmount

	expenseAmount = expense[index][1].value
	if expenseAmount != None:
		expenseAmount = 12 * expenseAmount
		yearlyExpense = yearlyExpense + expenseAmount


	yearlyNet = yearlyIncome - yearlyExpense
	
	if yearlyNet < 0.0:
		yearlyNet = yearlyNet / taxRate
	
	balanceAmount = balanceAmount + yearlyNet


	balance.append([balanceAmount, yearlyNet])


wb.close()



wb = load_workbook(excelFile)
sheet = wb.worksheets[0]


for rowIndex, row in enumerate(balance):
	for columnIndex, column in enumerate(row):
		sheet.cell(row = 39 + rowIndex, column = 8 + columnIndex).value = column
	print(row)


wb.save(excelFile)



