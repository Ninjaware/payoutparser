from transaction import Transaction
from datetime import datetime

COL_SRC_CHARGEDDATE = 1
COL_SRC_ORDERNO = 0
COL_SRC_CURRENCY = 10
COL_SRC_SALESINCLVAT = 13
COL_SRC_SALESEXCLVAT = 11
COL_SRC_BALANCE = 16
COL_SRC_FXRATE = 15
COL_SRC_VAT = 12
COL_SRC_PAYOUTDATE = 4

dateformat = '%Y-%m-%d'

def parsepayout(csvfile):
	
	tadic = {}
	
	# Move the cursor to the first csv row
	emptyrows = 0
	while csvfile.readline().startswith("Order Number") == False and emptyrows < 10:
		emptyrows = emptyrows + 1
	
	if(emptyrows == 10):
		print "Not a valid csv file:", csvfile.name
		return tadic

	# Loop the csv rows and add the transactions into a dictionary, where currency is the key
	for line in csvfile:
		row = line.split(',')
		ta = Transaction()
		ta.chargeddate = datetime.strptime(row[COL_SRC_CHARGEDDATE], dateformat)
		ta.orderno = row[COL_SRC_ORDERNO]
		ta.currency = row[COL_SRC_CURRENCY]
		ta.salesinclvat = float(row[COL_SRC_SALESINCLVAT])
		ta.salesexclvat = float(row[COL_SRC_SALESEXCLVAT])
		ta.balance = float(row[COL_SRC_BALANCE])
		if(ta.currency == 'EUR'):
			ta.fxrate = 1.00
		else:
			ta.fxrate = float(row[COL_SRC_FXRATE])
		ta.vat = float(row[COL_SRC_VAT])
		ta.payoutdate = datetime.strptime(row[COL_SRC_PAYOUTDATE], dateformat)
		
		if(tadic.has_key(ta.currency)):
			talist = tadic[ta.currency]
		else:
			talist = tadic[ta.currency] = []
		
		talist.append(ta)

	return tadic
