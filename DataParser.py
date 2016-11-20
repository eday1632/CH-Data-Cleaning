
import pandas as pd
import numpy as np
from os import listdir
from os.path import isfile, join
from shutil import move
import dateutil.parser as dparser
import datetime

pd.set_option('display.max_columns', 100)

def parseData(df, dftype=-1):
	
	if dftype == 0:
		return bgParser(df)

	elif dftype == 1:
		saksStylesParser(df)	

	elif dftype == 2:
		saksStoresParser(df)

	elif dftype == 3:
		saksSpecialsParser(df)

	elif dftype == 4:
		saksFiscalParser(df)

	elif dftype == 5:
		neimanStylesParser(df)

	elif dftype == 6:
		neimanStoresParser(df)

	elif dftype == 7:
		neimanOnlineParser(df)

	elif dftype == 8:
		neimanDashboardParser(df)

	else:
		print dftype + ' is out of range.'
		raise IndexError

def xldate_as_datetime(xldate):
    return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(xldate)))

def halt(df):
	print df.head(12)
	print df.shape
	raise Exception

def bgParser(df):
	try:
		refPoint = np.where(df == 'WEEKLY DETAILED SELLING')
		refRow = refPoint[0][0] + 1
		refCol = refPoint[1][0]
		date = xldate_as_datetime(df.loc[refRow][refCol]).date()
	except TypeError:
		refPoint = np.where(df[0] == 'Carolina Herrera ')
		refRow = refPoint[0][0] + 1
		refCol = 0
		date = dparser.parse(df.loc[refRow][refCol]).date()
	except IndexError:
		refPoint = np.where(df == 'CAROLINA HERRERA')
		refRow = refPoint[0][0] + 1
		refCol = refPoint[1][0]
		date = dparser.parse(df.loc[refRow][refCol], fuzzy=True).date()

	firstRow = np.where(df[0] == 'SN')[0][0]
	df = df.ix[firstRow+1:]

	dropCols = []
	namesList = []
	if df.shape[1] == 31:
		dropCols = [2,3,6,7,28,29,30]
		namesList = ['Season_Code','Group_Code','Sub_Dept','Product_Name','Color_Desc','Orig_Retail_Price','First_Receipt','Last_Receipt',\
		'Sales_U_WTD','Sales_$_WTD','Sales_U_MTD','Sales_$_MTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD',\
		'ST_U_LTD','ST_$_LTD','OH_U','OH_$','OO_U','OO_$']
	elif df.shape[1] == 29:
		dropCols = [2,3,6,7,27,28]
		namesList = ['Season_Code','Group_Code','Sub_Dept','Product_Name','Color_Desc','First_Receipt','Last_Receipt',\
		'Sales_U_WTD','Sales_$_WTD','Sales_U_MTD','Sales_$_MTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD',\
		'ST_U_LTD','ST_$_LTD','OH_U','OH_$','OO_U','OO_$']
	else:
		dropCols = [2,3,6,7,28,29]
		namesList = ['Season_Code','Group_Code','Sub_Dept','Product_Name','Color_Desc','Orig_Retail_Price','First_Receipt','Last_Receipt',\
		'Sales_U_WTD','Sales_$_WTD','Sales_U_MTD','Sales_$_MTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD',\
		'ST_U_LTD','ST_$_LTD','OH_U','OH_$','OO_U','OO_$']
		

	df = df.drop(df.columns[dropCols], axis=1)
	df.columns = namesList

	df = df.dropna(subset=['Product_Name'])
	df = df.reset_index(drop=True)

	df['Store'] = 'bg'
	df['Report_Date'] = date
	df['Channel'] = 'total'

	return df

def saksStylesParser(df):
	tempdf = df
	try:
		try:
			date = dparser.parse(df.loc[0][0], fuzzy=True).date()
		except AttributeError:
			date = -1
			pass

		df = df.dropna(subset=[5])
		df = df.reset_index(drop=True)

		df = df.drop(df.columns[[0,2,3,4,6,7,8,10,11,36]],axis=1)

		df.columns = ['Season_Code','Product_Name','Color_Desc','First_Receipt','Last_Receipt','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD',
		'Sales_U_PTD','Sales_$_PTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$','Rcpt_U_STD',
		'Rcpt_$_STD','Rcpt_U_LTD','Rcpt_$_LTD','ST_$_WTD','ST_$_PTD','ST_$_STD','ST_$_LTD','Cost']

		df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
		df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)

		if date == -1:
			date = df.Last_Receipt.max()

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except IndexError:
		return saksStylesParser1(tempdf)

def saksStylesParser1(df):
	tempdf = df
	try:
		if 'Figures' in df.loc[0][0]:
			date = dparser.parse(df.loc[0][0], fuzzy=True).date()
		else:
			date = dparser.parse(df.loc[5][0], fuzzy=True).date()

		df = df.drop(df.columns[[0,2,4,6,8]],axis=1)
		df.columns = ['Season_Code','Product_Name','Sub_Dept','Color_Desc','First_Receipt','Last_Receipt','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD',
			'Sales_U_PTD','Sales_$_PTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$','Rcpt_U_STD',
			'Rcpt_$_STD','Rcpt_U_LTD','Rcpt_$_LTD','ST_$_WTD','ST_$_PTD','ST_$_STD','ST_$_LTD']

		df = df.dropna(subset=['Product_Name'])
		df = df.reset_index(drop=True)

		df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
		df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except IndexError:
		return saksStylesParser2(tempdf)
	except ValueError:
		return saksStylesParser7(tempdf)

def saksStylesParser2(df):
	try:
		date = dparser.parse(df.loc[4][24], fuzzy=True).date()

		df = df.drop(df.columns[[0,1,2,4]],axis=1)
		df = df.dropna(subset=[3, 5])
		df = df.reset_index(drop=True)

		df.columns = ['Season_Code','Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','First_Receipt','Last_Receipt','Sales_U_WTD','Sales_$_WTD',
			'Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD','OH_U','OH_$','OO_U','OO_$','ST_$_WTD','ST_$_STD','ST_$_LTD']

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except AttributeError:
		return saksStylesParser3(df)

def saksStylesParser3(df):
	try:
		date = dparser.parse(df.loc[2][5], fuzzy=True).date()

		df = df.drop(df.columns[[0,1,2,4]],axis=1)
		df = df.dropna(subset=[3, 18])
		df = df.reset_index(drop=True)

		df.columns = ['Season_Code','Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','First_Receipt','Last_Receipt','Sales_U_WTD','Sales_$_WTD',
			'Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD','OH_U','OH_$','OO_U','OO_$','ST_$_WTD','ST_$_STD','ST_$_LTD']

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except AttributeError:
		return saksStylesParser4(df)

def saksStylesParser4(df):
	tempdf = df
	try:
		df = df.dropna(subset=[1])
		df = df.ix[13:]
		df = df.reset_index(drop=True)
		df = df.drop(df.columns[[1,2,3,4,5,7,9,11,14]], axis=1)
		df.columns = ['Season_Code','Product_Name','Sub_Dept','Color_Desc','First_Receipt','Last_Receipt','Sales_$_WTD','Sales_U_WTD','ST_$_WTD','Sales_$_PTD','Sales_U_PTD',
		'ST_$_PTD','Sales_$_STD','Sales_U_STD','ST_$_STD','Sales_$_LTD','Sales_U_LTD','ST_$_LTD','Rcpt_$_STD','Rcpt_U_STD','Rcpt_$_LTD','Rcpt_U_LTD','OH_$','OH_U',
		'OO_$','OO_U']

		df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
		df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)

		date = df.Last_Receipt.max().date()

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except ValueError:
		return saksStylesParser5(tempdf)

def saksStylesParser5(df):
	tempdf = df
	try:
		df = df.dropna(subset=[4])
		df = df.reset_index(drop=True)
		df = df.drop(df.columns[[1,3,5,10]], axis=1)

		df.columns = ['Season_Code','Product_Name','Sub_Dept','Color_Desc','First_Receipt','Last_Receipt','Sales_U_WTD',
		'Sales_$_WTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$','Rcpt_U_LTD','Rcpt_$_LTD',
		'ST_$_WTD','ST_$_STD','ST_$_LTD']

		df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
		df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)

		date = df.Last_Receipt.max().date()

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except ValueError:
		return saksStylesParser6(tempdf)

def saksStylesParser6(df):
	reportDate = dparser.parse(df.loc[0][0], fuzzy=True).date()
	df = df.drop(df.columns[[0,1,3,4,5,6,8,10,11,12,13,15,16,17]], axis=1)

	df.columns = ['Season_Code','Sub_Dept','Product_Name','Color_Desc','First_Receipt','Last_Receipt','Sales_$_LTD','Sales_U_LTD','Rcpt_$_LTD','Rcpt_U_LTD',
	'ST_$_LTD','OH_$','OH_U','OO_$','OO_U']

	df = df.dropna(subset=['Season_Code','Sub_Dept'])
	df = df.reset_index(drop=True)

	df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
	df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)

	df['Store'] = 'sfa'
	df['Report_Date'] = reportDate
	df['Type'] = 'Regular'
	
	return df

def saksStylesParser7(df):
	reportDate = dparser.parse(df.loc[0][0], fuzzy=True).date()

	df = df.drop(df.columns[[0,2,3,4,6,7,8,15,16,17,18,25,26,29,30,31,32]],axis=1)

	df.columns = ['Season_Code','Product_Name','Color_Desc','First_Receipt','Last_Receipt','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$','Rcpt_U_LTD','Rcpt_$_LTD']

	df['Product_Name'] = df['Product_Name'].fillna(method='ffill')
	df['Season_Code'] = df['Season_Code'].fillna(method='ffill')
	df = df.dropna(subset=['Product_Name','First_Receipt'])

	df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)
	df.Last_Receipt = df.Last_Receipt.apply(xldate_as_datetime)
	df['Sales_$_LTD'] = df['Sales_$_LTD'].astype(float)/1000
	df['Rcpt_$_LTD'] = df['Rcpt_$_LTD'].astype(float)/1000
	df['Sales_$_WTD'] = df['Sales_$_WTD'].astype(float)/1000
	df['OH_$'] = df['OH_$'].astype(float)/1000
	df['OO_$'] = df['OO_$'].astype(float)/1000

	df['Store'] = 'sfa'
	df['Report_Date'] = reportDate
	df['Type'] = 'Regular'

	return df

def saksStoresParser(df):
	tempdf = df
	try:
		firstRow = np.where(df[0] == 'Dept')[0][0]
		assert ('Day' in df.loc[2][0]), 'SeasonError: not Day collection'

		text = df.loc[0][0]
		date = dparser.parse(text, fuzzy=True).date()

		df = df.ix[firstRow+1:]
		df = df.drop(df.columns[[0,1]], axis=1)
		df.columns = ['City_Code','City','Sales_U_STD','Sales_$_STD','ST_$_STD','Sales_U_LTD','Sales_$_LTD','ST_$_LTD','OO_U','OO_$','OH_U','OH_$','Rcpt_U_LTD','Rcpt_$_LTD']
		
		df = df.dropna(subset=['City'])
		df = df.reset_index(drop=True)

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Season_Code'] = 'Day'
		df['Type'] = 'Regular'

		return df
	except TypeError:
		return saksStoresParser1(tempdf)

def saksStoresParser1(df):
	try:	
		firstRow = np.where(df[2] == 'NEW Store #')[0][0]
		seasons = np.where(df[2] == 'NEW Store #')[0]
		
		text = df.loc[1][22]
		date = xldate_as_datetime(text).date()

		if(df.shape[1] < 24):
			df['junk'] = np.nan
		df = df.drop(df.columns[[0,1,4,13,14,17,18,19,20,21,22,23]], axis=1)

		df['Season_Code'] = np.nan
		for season in seasons:
			df.ix[season, 'Season_Code'] = df.loc[season-1][4]
		df.Season_Code.fillna(method='ffill', inplace=True)

		df = df.ix[firstRow+1:]
		df.columns = ['City_Code','City','Sales_$_LTD','ST_$_LTD','Sales_$_LTD_LY','ST_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','OH_$','OH_$_LY','OO_$','OO_$_LY','Season_Code']
		
		df = df.dropna(subset=['City_Code'])
		df = df.dropna(subset=['City'])
		df = df.reset_index(drop=True)

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Regular'

		return df
	except TypeError:
		return saksStoresParser2(df)

def saksStoresParser2(df):
	salesType = ''
	if df.loc[2][4] == 'D.376':
		salesType = 'Regular'
	else: 
		salesType = 'Specials'

	markers = np.where(df[3] == 'Store')[0]
	firstRow = markers[0]
	date = xldate_as_datetime(df.loc[1][28])

	df['Season_Code'] = np.nan
	for marker in markers:
		df.ix[marker, 'Season_Code'] = df.loc[marker-2][4]
	df.Season_Code.fillna(method='ffill', inplace=True)

	df = df.ix[firstRow+1:]
	df = df[[3,4,5,6,8,9,10,15,16,18,19,20,'Season_Code']]

	df.columns = ['City_Code','Sales_Rank','Sales_$_LTD','ST_$_LTD','Rcpt_$_LTD','OH_$','OO_$','Sales_$_LTD_LY', \
	'ST_$_LTD_LY','Rcpt_$_LTD_LY','OH_$_LY','OO_$_LY','Season_Code']
	df.Sales_Rank.replace(to_replace='Sls Rank', value=np.nan, inplace=True)
	
	df = df.dropna(subset=['Sales_Rank','City_Code'])
	df = df.reset_index(drop=True)

	df['Store'] = 'sfa'
	df['Report_Date'] = date
	df['Type'] = salesType

	return df

def saksSpecialsParser(df):
	try:
		firstRow = np.where(df == 'Vendor')[0][0]

		text = df.loc[0][0]
		date = dparser.parse(text, fuzzy=True).date()

		df = df.ix[firstRow+1:]
		df = df.drop(df.columns[[0,1,2,3,4,26,27]], axis=1)
		df.columns = ['Season_Code','City_Code','City','Sales_U_WTD','Sales_$_WTD','Sales_U_STD','Sales_$_STD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$',
		'Rcpt_U_STD','Rcpt_$_STD','Rcpt_U_LTD','Rcpt_$_LTD','ST_$_WTD','ST_$_STD','ST_$_LTD','Cost']
		
		df = df.dropna(subset=['City'])
		df = df.reset_index(drop=True)

		df['Store'] = 'sfa'
		df['Report_Date'] = date
		df['Type'] = 'Specials'

		return df
	except TypeError:
		return saksSpecialsParser1(df)

def saksSpecialsParser1(df):
	firstRow = np.where(df[2] == 'NEW Store #')[0][0]
	seasons = np.where(df[2] == 'NEW Store #')[0]
	
	text = df.loc[1][22]
	date = xldate_as_datetime(text).date()

	df['Season_Code'] = np.nan
	for season in seasons:
		df.ix[season, 'Season_Code'] = df.loc[season-1][3]
	df.Season_Code.fillna(method='ffill', inplace=True)

	df = df.drop(df.columns[[0,1,4,13,14,17,18,19,20,21,22]], axis=1)
	df = df.ix[firstRow+1:]
	df.columns = ['City_Code','City','Sales_$_LTD','ST_$_LTD','Sales_$_LTD_LY','ST_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','OH_$','OH_$_LY','OO_$','OO_$_LY','Season_Code']
	
	df = df.dropna(subset=['City_Code'])
	df = df.dropna(subset=['City'])
	df = df.reset_index(drop=True)

	df['Store'] = 'sfa'
	df['Report_Date'] = date
	df['Type'] = 'Specials'

	return df

def saksFiscalParser(df):
	regular = '376 - HERRERA /CH'
	specials = '527 - CH   HERRERA SPECIAL ORDE'
	if df.loc[6][0] == regular:
		return saksFiscalStoreParser(df, regular)
	elif df.loc[6][0] == specials:
		return saksFiscalStoreParser(df, specials)
	else:
		text = df.loc[9][1]
		date = dparser.parse(text, fuzzy=True).date()

		df[0].fillna(method='ffill', limit=1, inplace=True)
		df[0].fillna(method='bfill', limit=1, inplace=True)
		df[0].replace(to_replace='376 - HERRERA /CH  ', value='Regular', inplace=True)
		df[0].replace(to_replace='527 - CH   HERRERA SPECIAL ORDE  ', value='Specials', inplace=True)

		df = df.dropna(subset=[1,2])
		df = df.drop(df.columns[[2,3,4,5,6,7,8,10,11,12,13,14,16,17,18,19,20,21,22,25,28]], axis=1)
		df.columns = ['Type','Channel','Plan_EOM','Plan_EOS','Sales_$_MTD','Sales_$_MTD_LY','Sales_$_STD','Sales_$_STD_LY']

		df['Store'] = 'sfa'
		df['Report_Date'] = date

		return df

def saksFiscalStoreParser(df, salesType):
	text = df.loc[4][26]
	date = dparser.parse(text, fuzzy=True).date()

	firstRow = np.where(df[0] == salesType)[0][0] + 4
	lastRow = np.where(df[0] == salesType)[0][1] - 2

	df = df.ix[firstRow:lastRow]

	df = df.drop(df.columns[[1,2,3,4,5,6,7,9,10,11,12,13,15,16,17,18,19,20,21,24,27,28]], axis=1)
	df.columns = ['City','Plan_EOM','Plan_EOS','Sales_$_MTD','Sales_$_MTD_LY','Sales_$_STD','Sales_$_STD_LY']

	df['Store'] = 'sfa'
	df['Report_Date'] = date

	df = df.reset_index(drop=True)

	return df

def neimanFiscalParser(df):


	halt(df)

def neimanStoresParser(df):
	try:	
		if df.loc[1][18] == 'Units':
			df = pd.DataFrame()
			return df
		try:
			firstRow = np.where(df[0] == 'TOTAL NEVA HALL')[0][0]
			lastRow = np.where(df[0] == 'TOTAL NEVA HALL')[0][1]
		except IndexError:
			return neimanStoresParser2(df)

		startDate = df.loc[2][18]
		endDate = df.loc[3][18]

		startDate = dparser.parse(startDate, fuzzy=True).date()
		endDate = dparser.parse(endDate, fuzzy=True).date()

		df = df.ix[firstRow:lastRow]
		df = df.drop(df.columns[[0,1,2,4,7,8,11,12,13,14,17,18,19]], axis=1)
		df.columns = ['City','OH_$','OH_$_LY','Sales_$_LTD','Sales_$_LTD_LY','ST_$_LTD','ST_$_LTD_LY']

		df = df.reset_index(drop=True)

		df['Rcpt_$_LTD'] = df['Sales_$_LTD'] + df['OH_$']
		df['Store'] = 'nm'
		df['Report_Date'] = endDate
		df['Start_Date'] = startDate

		return df
	except AttributeError:
		return neimanStoresParser1(df)

def neimanStoresParser1(df):
	if df.loc[1][17] == 'Units':
		df = pd.DataFrame()
		return df
	try:
		firstRow = np.where(df[0] == 'TOTAL NEVA HALL')[0][0]
		lastRow = np.where(df[0] == 'TOTAL NEVA HALL')[0][1]

		startDate = df.loc[2][17]
		endDate = df.loc[3][17]

		startDate = dparser.parse(startDate, fuzzy=True).date()
		endDate = dparser.parse(endDate, fuzzy=True).date()

		df = df.ix[firstRow:lastRow]
		df = df.drop(df.columns[[0,1,2,6,7,10,11,12,13,16,17,18]], axis=1)
		df.columns = ['City','OH_$','OH_$_LY','Sales_$_LTD','Sales_$_LTD_LY','ST_$_LTD','ST_$_LTD_LY']

		df = df.reset_index(drop=True)

		df['Rcpt_$_LTD'] = df['Sales_$_LTD'] + df['OH_$']
		df['Store'] = 'nm'
		df['Report_Date'] = endDate
		df['Start_Date'] = startDate

		return df
	except IndexError:
		return neimanStoresParser2(df)


def neimanStoresParser2(df):
	try:
		startDate = dparser.parse(df.loc[2][18], fuzzy=True).date()
		endDate = dparser.parse(df.loc[3][18], fuzzy=True).date()

		df = df.dropna(subset=[2])
		df = df.drop(df.columns[[0,1,2,4,7,8,11,12,13,14,17,18]], axis=1)
		df.columns = ['City_Code','OH_$','OH_$_LY','Sales_$_PTD','Sales_$_PTD_LY','ST_$_PTD','ST_$_PTD_LY','OO_$']
		df = df.reset_index(drop=True)

		df['Rcpt_$_PTD'] = df['Sales_$_PTD'] + df['OH_$']
		df['Store'] = 'nm'
		df['Report_Date'] = endDate
		df['Start_Date'] = startDate

		return df
	except AttributeError:
		return neimanStoresParser3(df)

def neimanStoresParser3(df):
	try:
		reportDate = df.loc[1][0]
		reportDate = dparser.parse(reportDate[:10], fuzzy=True).date()
		markers = np.where(df[2] == 'TOTAL')[0]

		df1 = pd.DataFrame()
		for marker in markers:
			tempdf = df.iloc[marker:marker+8].T
			tempdf.columns = [0,1,2,3,4,5,6,7]
			df1 = df1.append(tempdf.reset_index(drop=True), ignore_index=True)

		df1 = df1[[0,2,3,4,5,7]]
		df1.columns = ['City_Code','Rcpt_$_LTD','OH_$','Sales_$_LTD','OO_$','ST_$_LTD']
		df1['Season_Code'] = np.nan
		df1['Report_Date'] = reportDate
		df1['Store'] = 'nm'

		markers = np.where(df1['City_Code'] == 'TOTAL')[0]
		for marker in markers:
			df1.loc[marker, 'City_Code'] = np.nan
			df1.loc[marker+1, 'City_Code'] = np.nan
			df1.loc[marker, 'Season_Code'] = df1.loc[marker-2][0]

		df1.Season_Code.fillna(method='ffill', inplace=True)
		df1 = df1.dropna(subset=['City_Code','Season_Code'])
		df1 = df1.reset_index(drop=True)
		return df1
	except TypeError:
		return neimanStoresParser4(df)

def neimanStoresParser4(df):
	reportType = df.loc[4][1]
	if reportType == 'Dollars':
		reportType = '$'
	else: reportType = 'U'

	reportDate = xldate_as_datetime(df.loc[4][1]).date()

	anchor = np.where(df[1] == 'STR')[0][0]
	index = pd.Series(df[1][anchor+1:])

	markers = list(df.ix[8][df.ix[8] == 'OH'].index)
	df1 = pd.DataFrame()
	for marker in markers:
		tempdf = pd.DataFrame(df[[1, marker, marker+1, marker+2]])
		tempdf.columns = [0,1,2,3]
		df1 = df1.append(tempdf.reset_index(drop=True), ignore_index=True)

	df1.columns = ['City','OH_$','Sales_$_LTD','ST_$_LTD']
	df1['Season_Code'] = np.nan
	df1['Report_Date'] = reportDate
	df1['Store'] = 'nm'

	markers = np.where(df1['OH_$'] == 'OH')[0]
	for marker in markers:
		df1.loc[marker, 'Season_Code'] = df1.loc[marker-1][1]

	df1.Season_Code.fillna(method='ffill', inplace=True)
	df1 = df1.dropna(subset=['City','OH_$'])
	df1 = df1.reset_index(drop=True)
	df1.to_csv('U:\Retail Reporting\MagicFolder\Retail\Out\\test.csv', index=False)

	return df1

def neimanStylesParser(df):
	try:	
		if (df.loc[0][8] == 'Markdown') | (df.loc[0][8] == 'Total'):
			df = pd.DataFrame()
			return df

		firstRow = np.where(df[1] == 'CD')[0][0]
		startPos = np.where(df[7] == 'WK Start:')[0][0]
		endPos = np.where(df[7] == 'WK End:')[0][0]
		startDate = df.loc[startPos][8]
		endDate = df.loc[endPos][8]

		startDate = dparser.parse(startDate, fuzzy=True).date()
		endDate = dparser.parse(endDate, fuzzy=True).date()

		df = df.ix[firstRow+1:]
		df = df.drop(df.columns[[0,1,2,3,4,5,7,8,9,10,12,15,23,24,25,28,29,30]], axis=1)
		df.columns = ['Product_Name','Season_Code','Color_Desc','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_PTD','Sales_$_PTD','OH_U','OH_$',
		'ST_$_LTD','OO_U','OO_$']

		df = df.reset_index(drop=True)

		df['Rcpt_U_LTD'] = df.Sales_U_PTD + df.OH_U
		df['Rcpt_$_LTD'] = df['Sales_$_PTD'] + df['OH_$']
		df['Store'] = 'nm'
		df['Report_Date'] = endDate
		df['Start_Date'] = startDate

		return df
	except IndexError:
		return neimanStylesParser1(df)

def neimanStylesParser1(df):
	try:	
		date = xldate_as_datetime(df.loc[5][1]).date()
		totCols = df.shape[1]

		if totCols > 100:
			df = df.dropna(thresh=3, axis=1)
			totCols = df.shape[1]
			
		dropCols = []
		if totCols == 23:
			dropCols = [0,2,7,9,10,18,19,22]
			df = df.drop(df.columns[dropCols], axis=1)
			df.columns = ['Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','Rcpt_U_LTD','OH_U','Sales_U_WTD','ST_U_WTD','Sales_U_STD','ST_U_STD',
			'Sales_U_LTD','ST_U_LTD','Rcpt_$_LTD','Sales_$_LTD']

		elif totCols == 26:
			dropCols = [0,2,7,9,10,21,22,25]
			df = df.drop(df.columns[dropCols], axis=1)
			df.columns = ['Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','Rcpt_U_LTD','OO_U','OH_U','Sales_U_WTD','ST_U_WTD','Sales_U_STD','ST_U_STD',
			'Sales_U_YTD','ST_U_YTD','Sales_U_LTD','ST_U_LTD','Rcpt_$_LTD','Sales_$_LTD']

		elif totCols == 24:
			dropCols = [0,2,8,9,19,20,23]
			df = df.drop(df.columns[dropCols], axis=1)
			df.columns = ['Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','Rcpt_U_LTD','OH_U','Sales_U_WTD','ST_U_WTD','Sales_U_STD','ST_U_STD',
			'Sales_U_YTD','ST_U_YTD','Sales_U_LTD','ST_U_LTD','OH_$','Sales_$_YTD']

		elif totCols == 19:
			dropCols = [0,2,7,9,10,16]
			df = df.drop(df.columns[dropCols], axis=1)
			df.columns = ['Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','Rcpt_U_LTD','OH_U','Sales_U_WTD','ST_U_WTD','Sales_U_LTD','ST_U_LTD',
			'Rcpt_$_LTD','Sales_$_LTD']

		else:
			dropCols = [0,2,7,9,10,18,19]
			df = df.drop(df.columns[dropCols], axis=1)
			df.columns = ['Product_Name','Sub_Dept','Color_Desc','Cost','Orig_Retail_Price','Rcpt_U_LTD','OH_U','Sales_U_WTD','ST_U_WTD','Sales_U_STD','ST_U_STD',
			'Sales_U_LTD','ST_U_LTD','Rcpt_$_LTD','Sales_$_LTD']

		df = df.dropna(subset=['Color_Desc','Product_Name'])
		df = df.reset_index(drop=True)

		df['Store'] = 'nm'
		df['Report_Date'] = date
		df['Channel'] = 'store'
		
		return df
	except ValueError:
		return neimanStylesParser2(df)

def neimanStylesParser2(df):
	reportDate = dparser.parse(df.loc[2][1]).date()

	df[3].replace(to_replace='ITEM #', value=np.nan, inplace=True)
	df = df.dropna(subset=[3])
	df = df.reset_index(drop=True)

	df = df[[5,6,8,9,11,12,13,15,16,18,19,20]]
	df.columns = ['Orig_Retail_Price','Product_Name','Sales_U_LTD','Rcpt_U_LTD','ST_U_LTD','OH_U','OO_U','Sales_$_LTD','Rcpt_$_LTD','ST_$_LTD','OH_$','OO_$']

	df['Store'] = 'nm'
	df['Report_Date'] = reportDate
	df['Channel'] = 'store'

	return df

def neimanOnlineParser(df):
	tempdf = df
	try:
		reportDate = -1
		try:
			if 'Image' in df.loc[0][1]:
				reportDate = 'unknown'
				df = df.ix[1:]
			else:
				reportDate = dparser.parse(df.loc[3,6]).date()
		except TypeError:
			pass		

		shape = df.shape[1]
		if shape == 109:
			df = df[[4,6,24,25,26,27,28,29,30,31,40,47,48,51,52,53,54,59,60,61,93,96,103]]
			df.columns = ['Sales_$_LTD','Sales_U_LTD','Sales_$_WTD','Sales_U_WTD','Sales_$_MTD','Sales_U_MTD','Sales_$_STD','Sales_U_STD','Sales_$_YTD','Sales_U_YTD',
			'ST_U_STD','OH_$','OH_U','Rcpt_$_LTD','Rcpt_U_LTD','Rcpt_$_STD','Rcpt_U_STD','OO_$','OO_U','First_Receipt','Product_Name','Season_Code','Orig_Retail_Price']
		elif shape == 108:
			df = df[[4,6,24,25,26,27,28,29,30,31,40,46,47,50,51,52,53,58,59,60,92,95,102]]
			df.columns = ['Sales_$_LTD','Sales_U_LTD','Sales_$_WTD','Sales_U_WTD','Sales_$_MTD','Sales_U_MTD','Sales_$_STD','Sales_U_STD','Sales_$_YTD','Sales_U_YTD',
			'ST_U_STD','OH_$','OH_U','Rcpt_$_LTD','Rcpt_U_LTD','Rcpt_$_STD','Rcpt_U_STD','OO_$','OO_U','First_Receipt','Product_Name','Season_Code','Orig_Retail_Price']
		else:
			df = df[[3,4,19,20,21,22,23,24,25,26,35,41,42,45,46,47,48,55,87,90,97]]
			df.columns = ['Sales_$_LTD','Sales_U_LTD','Sales_$_WTD','Sales_U_WTD','Sales_$_MTD','Sales_U_MTD','Sales_$_STD','Sales_U_STD','Sales_$_YTD','Sales_U_YTD',
			'ST_U_STD','OH_$','OH_U','Rcpt_$_LTD','Rcpt_U_LTD','Rcpt_$_STD','Rcpt_U_STD','First_Receipt','Product_Name','Season_Code','Orig_Retail_Price']

		df = df.dropna(subset=['First_Receipt','Season_Code','Sales_U_LTD'])
		df = df.reset_index(drop=True)

		df.First_Receipt = df.First_Receipt.apply(xldate_as_datetime)

		df['Store'] = 'nm'
		df['Report_Date'] = reportDate
		df['Channel'] = 'online'

		return df
	except KeyError:
		return neimanOnlineParser1(tempdf)
	except AttributeError:
		return neimanOnlineParser2(tempdf)

def neimanOnlineParser1(df):
	try:
		reportDate = dparser.parse(df.loc[2][7]).date()

		df = df[[2,6,10,11,12,13,14,15,16,17,18,32,33,36,37,38,39,44,45]]
		df.columns = ['Product_Name','Season_Code','Sales_$_WTD','Sales_U_WTD','Sales_$_STD','Sales_U_STD','ST_$_STD','Sales_$_YTD','Sales_U_YTD','Sales_$_LTD',
		'Sales_U_LTD','OH_$','OH_U','Rcpt_$_LTD','Rcpt_U_LTD','Rcpt_$_STD','Rcpt_U_STD','OO_$','OO_U']

		df['Store'] = 'nm'
		df['Report_Date'] = reportDate
		df['Channel'] = 'online'
		df = df.dropna(subset=['Product_Name'])
		df = df.reset_index(drop=True)

		return df
	except ValueError:
		return neimanOnlineParser2(df, reroute=False)

def neimanOnlineParser2(df, reroute=True):
	tempdf = df
	try:
		reportDate = xldate_as_datetime(df.loc[3][8])

		df = df[[2,6,7,9,13,14,15,16,17,18,19,20,21,43,44,45,46,47,48,53,54]]
		df.columns = ['Product_Name','Color_Desc','Size_Code','Season_Code','Sales_$_WTD','Sales_U_WTD','Sales_$_STD','Sales_U_STD','ST_$_STD','Sales_$_YTD',
		'Sales_U_YTD','Sales_$_LTD','Sales_U_LTD','OH_$','OH_U','Rcpt_$_LTD','Rcpt_U_LTD','Rcpt_$_STD','Rcpt_U_STD','OO_$','OO_U']

		df['Store'] = 'nm'
		df['Report_Date'] = reportDate
		df['Channel'] = 'online'

		df.Product_Name.replace(to_replace='Ven Style Ref', value=np.nan, inplace=True)
		df = df.dropna(subset=['Product_Name'])
		df = df.reset_index(drop=True)

		return df
	except ValueError:
		if kill:
			return neimanOnlineParser1(tempdf)

def neimanDashboardParser(df):
	try:
		reportDate = dparser.parse(df.loc[4][7]).date()
	except AttributeError:
		try:
			reportDate = dparser.parse(df.loc[4][8]).date()
		except AttributeError:
			reportDate = dparser.parse(df.loc[4][9]).date()

	period = ''
	if 'YTD' in df.loc[1][1]:
		period = 'YTD'
	else:
		period = 'STD'

	df = df.dropna(thresh=10, axis=1)
	df = df[[1,3,4]]
	df.columns = ['Fiscal_Category', 'Fiscal_Sales','Fiscal_Sales_LY']

	df['Period'] = period
	df['Store'] = 'nm'
	df['Report_Date'] = reportDate
	df['Channel'] = 'online'
	
	df = df.dropna(subset=['Fiscal_Category','Fiscal_Sales'])
	df = df.reset_index(drop=True)

	return df

