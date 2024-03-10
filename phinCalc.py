#import libraries 
from openpyxl import load_workbook


excelFile = 'phinCalc.xlsx'


wb = load_workbook(excelFile)
sheet = wb.worksheets[0]


initialInvestment = sheet['B5':'D9']
income = sheet['B16':'H25']
expense = sheet['B32':'H41']
otherInformation = sheet['C46':'D50']


balanceAmount = 0.0

for row in initialInvestment:
	cellValue = row[2].value
	if cellValue is not None:
		balanceAmount = balanceAmount + cellValue

print(balanceAmount)


startYear = otherInformation[0][1].value
startAge = otherInformation[1][1].value
endAge = otherInformation[2][1].value
interestRateOnBalance = otherInformation[3][1].value/100.0 + 1
taxRate = 1 - otherInformation[4][1].value/100.0

balance = []
# balance.append([startYear, balance])
# print(balance)

for index, age in enumerate(range(startAge, endAge + 1)):

	year = startYear + index


	balanceAmount = balanceAmount * interestRateOnBalance


	yearlyIncome = 0.0
	
	for row in income:
		incomeAmount = row[2].value
		if incomeAmount == None:
			continue

		incomeYear = row[3].value
		if incomeYear == None:
			continue

		if year == incomeYear:
			yearlyIncome = yearlyIncome + incomeAmount

	for row in income:
		incomeAmount = row[4].value
		if incomeAmount == None:
			continue
		incomeAmount = 12 * incomeAmount

		incomeStartYear = row[5].value
		if incomeStartYear == None:
			continue

		incomeEndYear = row[6].value
		if incomeEndYear == None:
			continue

		if year >= incomeStartYear and year <= incomeEndYear:
			yearlyIncome = yearlyIncome + incomeAmount

	yearlyIncome = taxRate * yearlyIncome


	yearlyExpense = 0.0
	
	for row in expense:
		expenseAmount = row[2].value
		if expenseAmount == None:
			continue

		expenseYear = row[3].value
		if expenseYear == None:
			continue

		if year == expenseYear:
			yearlyExpense = yearlyExpense + expenseAmount

	for row in expense:
		expenseAmount = row[4].value
		if expenseAmount == None:
			continue
		expenseAmount = 12 * expenseAmount

		expenseStartYear = row[5].value
		if expenseStartYear == None:
			continue

		expenseEndYear = row[6].value
		if expenseEndYear == None:
			continue

		if year >= expenseStartYear and year <= expenseEndYear:
			yearlyExpense = yearlyExpense + expenseAmount


	yearlyNet = yearlyIncome - yearlyExpense
	
	if yearlyNet < 0:
		yearlyNet = yearlyNet / taxRate
	
	balanceAmount = balanceAmount + yearlyNet


	balance.append([age, year, balanceAmount, yearlyNet])


for rowIndex, row in enumerate(balance):
	for columnIndex, column in enumerate(row):
		sheet.cell(row = 56 + rowIndex, column = 4 + columnIndex).value = column
	print(row)


wb.save(excelFile)





