import openpyxl
import sys
import collections

filename = sys.argv[1]
print('opening:  %s' % filename)
wb = openpyxl.load_workbook(filename)
sheet = wb.get_sheet_by_name('Sheet1')

Invoice = collections.namedtuple('Invoice', ['date', 'id', 'po', 'job', 'amount'])
invoices = []
for row in range(3, sheet.get_highest_row()):
	invoices.append(Invoice('1', '2', '3', '4'))

print(invoices)	