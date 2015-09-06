import openpyxl
import sys
import collections
import pytest
import os

Invoice = collections.namedtuple('Invoice', ['date', 'id', 'po', 'job', 'amount'])
Col_Map = collections.namedtuple('Col_Map', ['date', 'id', 'po', 'name', 'job', 'amount'])
xl_map = Col_Map(3, 4, 5, 6, 9, 10)
sheet_title = 'Sheet1'
start_row = 3
company = 'Next Door Distribution Company'

def get_invoices(filename):
	print('opening:  %s' % filename)
	wb = openpyxl.load_workbook(filename)
	sheet = wb.get_sheet_by_name(sheet_title)
	
	invoices = []
	for row in range(start_row, sheet.get_highest_row() + 1):
		invoices.append(Invoice('1', '2', '3', '4', '5'))

	return(invoices)

@pytest.fixture
def filename(request):
	fname = 'xtest.xlsx'
	wb = openpyxl.Workbook()
	sheet = wb.get_active_sheet()
	sheet.title = sheet_title
	sheet.cell(row=start_row, column=xl_map.date).value = '09/06/2015'
	sheet.cell(row=start_row, column=xl_map.id).value = 12345
	sheet.cell(row=start_row, column=xl_map.po).value = 'FL1007-010'
	sheet.cell(row=start_row, column=xl_map.name).value = company
	sheet.cell(row=start_row, column=xl_map.job).value = '1007.R57'
	sheet.cell(row=start_row, column=xl_map.amount).value = 3400.75
	wb.save(fname)

	def fin():
		os.remove(fname)
	request.addfinalizer(fin)
	return fname

def test_rdr(filename):
	result = get_invoices(filename)

	assert len(result) == 1
