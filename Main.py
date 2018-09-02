# Xavier Collantes
# 09/02/18
# Examples tried out from Al Sweigart's book, "Automating the Boring Stuff with Python"
# Excel Automation: Chapter 12

import openpyxl


def t(obj):
	print (type(obj))
	
	
if __name__ == '__main__':
	xlsx = 'xl/example.xlsx'
	wb = openpyxl.load_workbook(xlsx)
	print(wb.sheetnames)
	fsheet = wb['Fruits']
	
	print(type(fsheet['A1'].value))
	print(type(fsheet['ZZ1'].value))
	fsheet['ZZ1'].value = 'X'
	print(fsheet['ZZ1'].value)
	
	m = fsheet['B1']
	print("Cell: " + str(m.coordinate))
	print("Row: " + str(m.row) + "|  Column: " + str(m.column))
	
	
	
	
	

	