from transaction import Transaction
from openpyxl.workbook import Workbook
from openpyxl.cell import get_column_letter
from openpyxl.style import NumberFormat

COL_DST_CHARGEDDATE = 0
COL_DST_ORDERNO = 1
COL_DST_CURRENCY = 2
COL_DST_SALESINCLVAT = 3
COL_DST_SALESEXCLVAT = 4
COL_DST_SERVICEFEE = 5
COL_DST_BALANCE = 6
COL_DST_VAT = 7
COL_DST_PAYOUTDATE = 8

header = {}
header[COL_DST_CHARGEDDATE] = 'Charged date'
header[COL_DST_ORDERNO] = 'Merchant order number'
header[COL_DST_CURRENCY] = 'Currency'
header[COL_DST_SALESINCLVAT] = 'Sales incl. VAT'
header[COL_DST_SALESEXCLVAT] = 'Sales excl. VAT'
header[COL_DST_SERVICEFEE] = 'Google service fee'
header[COL_DST_BALANCE] = 'Balance'
header[COL_DST_VAT] = 'VAT'
header[COL_DST_PAYOUTDATE] = 'Payout date'

sumcols = [COL_DST_SALESINCLVAT, COL_DST_SALESEXCLVAT, COL_DST_SERVICEFEE, COL_DST_BALANCE, COL_DST_VAT]

def saveasworkbook(tadic, outputfilename):
	
	workbook = Workbook()
	sheet = workbook.worksheets[0]

	row = 0 # cursor for the output row
	for currency, talist in tadic.iteritems():
		row += createheader(sheet, row)
		grouprows = creategroup(sheet, row, currency, talist)
		row += grouprows
		row += createtotals(sheet, row, grouprows, currency)
		row += 1 # separator row
	
	# set column widths manually
	sheet.column_dimensions[get_column_letter(COL_DST_CHARGEDDATE+1)].width = 12
	sheet.column_dimensions[get_column_letter(COL_DST_ORDERNO+1)].width = 20
	sheet.column_dimensions[get_column_letter(COL_DST_CURRENCY+1)].width = 9
	sheet.column_dimensions[get_column_letter(COL_DST_SALESINCLVAT+1)].width = 13
	sheet.column_dimensions[get_column_letter(COL_DST_SALESEXCLVAT+1)].width = 13
	sheet.column_dimensions[get_column_letter(COL_DST_SERVICEFEE+1)].width = 15
	sheet.column_dimensions[get_column_letter(COL_DST_BALANCE+1)].width = 7
	sheet.column_dimensions[get_column_letter(COL_DST_VAT+1)].width = 6
	sheet.column_dimensions[get_column_letter(COL_DST_PAYOUTDATE+1)].width = 12

	# save the workbook
	workbook.save(outputfilename)

# writes the transaction values and returns the number of rows in the group
def creategroup(sheet, row, currency, talist):
	for ta in talist:
		sheet.cell(row=row, column=COL_DST_CHARGEDDATE).value = ta.chargeddate
		sheet.cell(row=row, column=COL_DST_ORDERNO).style.number_format.format_code = NumberFormat.FORMAT_TEXT
		sheet.cell(row=row, column=COL_DST_ORDERNO).value = ta.orderno
		sheet.cell(row=row, column=COL_DST_CURRENCY).value = ta.currency
		sheet.cell(row=row, column=COL_DST_SALESINCLVAT).value = ta.salesinclvat
		sheet.cell(row=row, column=COL_DST_SALESEXCLVAT).value = ta.salesexclvat
		sheet.cell(row=row, column=COL_DST_SERVICEFEE).value = ta.servicefee()
		sheet.cell(row=row, column=COL_DST_BALANCE).value = ta.balance
		sheet.cell(row=row, column=COL_DST_VAT).value = ta.vat
		sheet.cell(row=row, column=COL_DST_PAYOUTDATE).value = ta.payoutdate
		row += 1

	return len(talist)

# writes the total formulas
def createtotals(sheet, row, grouprows, currency):
	sheet.cell(row=row, column=COL_DST_CURRENCY).value = 'Total ' + currency
	sheet.cell(row=row, column=COL_DST_CURRENCY).style.font.bold = True	
	for sumcol in sumcols:
		colletter = get_column_letter(sumcol+1)
		cell = sheet.cell(row=row, column=sumcol)
		cell.value = '=SUM(' + colletter + str(row-grouprows+1) + ':' + colletter + str(row) + ')'
		cell.style.font.bold = True

	return 1

# writes the header captions
def createheader(sheet, row):
	for column, text in header.iteritems():
		sheet.cell(row=row, column=column).value = text
	return 1

