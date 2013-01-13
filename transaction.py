
class Transaction:
	chargeddate = 0
	orderno = ""
	currency = ""
	salesinclvat = 0.00
	salesexclvat = 0.00
	balance = 0.00
	fxrate = 0.00
	vat = 0.00
	payoutdate = 0

	def servicefee(self):
		return self.fxrate * self.salesinclvat - self.balance
