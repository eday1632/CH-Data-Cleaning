
from os import listdir
from os.path import isfile, join
import xlrd, csv
from shutil import move
import Tkinter, tkMessageBox

def convertExcelToCSV(dir, filename):
	try:
		workbook = xlrd.open_workbook(join(directory, filename))
		all_worksheets = workbook.sheet_names()
		tempFile = stripFileExtension(filename)
		i = 0
		for worksheet_name in all_worksheets:
			worksheet = workbook.sheet_by_name(worksheet_name)
			your_csv_file = open(join('C:\Users\eday\Google Drive\Work\CalFresh\Data\DFA296X\\', tempFile+'-'+str(i)+'.csv'), 'wb')
			i = i + 1
			wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

			for rownum in xrange(worksheet.nrows):
			    wr.writerow([unicode(entry).encode("utf-8") for entry in worksheet.row_values(rownum)])
			your_csv_file.close()
	except Exception as e:
		print e

def stripFileExtension(filename):
	if filename.endswith('.xlsx'):
		return filename[:-5]
	else:
		return filename[:-4]

def dealWithFile(filename):
	if tkMessageBox.askyesno('Unfamiliar file!', 'Should we delete it? \n\n'+filename):
		print 'Poof! The file is gone.'
		move(join(directory, filename), 'C:\Users\eday\Google Drive\Work\CalFresh\Junk\\'+filename)
	else:
		print 'Then tell me how to identify it!'
		raise Exception



directory = 'C:\Users\eday\Google Drive\Work\CalFresh\Data\DFA296X\\'
rawData = [f for f in listdir(directory) if isfile(join(directory, f))]

for f in rawData:
	print '\n', f
	# loop through the excel files to convert each to csv
	convertExcelToCSV(directory, f)


	