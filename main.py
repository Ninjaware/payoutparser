#!/usr/bin/python

import sys
from payoutparser import parsepayout
from sheetcreator import saveasworkbook

if len(sys.argv) < 2:
	print "Usage: payoutparser <csv file> [<output.xlsx>]"
	exit()
elif len(sys.argv) >= 3:
	outputfilename = sys.argv[2]
else:
	outputfilename = 'output.xlsx'

csvfilename = sys.argv[1]

try:
	csvfile = open(csvfilename, 'r')
except IOError:
	print "No such file:", csvfilename
	exit()

tadic = parsepayout(csvfile)
csvfile.close()

if len(tadic) > 0:
	saveasworkbook(tadic, outputfilename)
