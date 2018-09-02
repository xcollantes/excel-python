# Xavier Collantes
# 09/02/18
# Examples tried out from Al Sweigart's book, "Automating the Boring Stuff with Python"
# Excel Automation: Chapter 12

import openpyxl, os


def t(obj):
	print (type(obj))
	
	
def newSheet():
	wb = openpyxl.Workbook()
	
	
	ws = wb.create_sheet("Anotha One")
	wb.create_sheet("Anotha Anotha One")

	
	
	wb.save('./xl/myWS.xlsx')
	
newSheet()