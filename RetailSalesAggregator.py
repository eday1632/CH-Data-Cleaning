# coding: utf-8

import pandas as pd
import numpy as np
import os, unicodedata
from datetime import datetime
from datetime import date
from BookingsProcessor import processBookings
from MasterFilter import processData, cleanColors, stateCleaner
from IconSalesAnalyzer import analyzeIconSales
import Tkinter, tkMessageBox

pd.set_option('display.max_columns', 100)
pd.options.mode.chained_assignment = None
PRODUCTION = False
EOMREPORT = False

if PRODUCTION: 
    Tkinter.Tk().withdraw()
    tkMessageBox.showinfo('It\'s running!','Retail Fox will take 1 to 2 minutes.')

print 'Away we go!'

attrMap = 'U:\Retail Reporting\AppFolder\Retail\In\AttrMap.xlsx'

oldData = 'U:\Retail Reporting\AppFolder\Retail\In\SalesData.xlsx'

clients = pd.read_csv('U:\Retail Reporting\AppFolder\Retail\In\Vip Purchase History Report.csv', encoding='utf-8', error_bad_lines=False,
    usecols=['Store_Name','VIPName','Country','transdate','Ext_Price_Sold','Price_sold','Qty_Sold','Prov'])

priceDict = pd.io.excel.read_excel(attrMap, 1)
classDict = pd.io.excel.read_excel(attrMap, 2)
descDict = pd.io.excel.read_excel(attrMap, 3)
dateDict = pd.io.excel.read_excel(attrMap, 4)
seasonDict = pd.io.excel.read_excel(attrMap, 7)

priceDict = dict(zip(priceDict['Key'].astype(str), priceDict['Value']))
classDict = dict(zip(classDict['Key'].astype(str), classDict['Value'].str.lower()))
descDict = dict(zip(descDict['Key'].astype(str), descDict['Value'].str.lower()))
dateDict = dict(zip(dateDict['Key'].astype(str), dateDict['Value']))
seasonDict = dict(zip(seasonDict['Key'].astype(str), seasonDict['Value']))

# print 'Domestic...'

# domestic = pd.io.excel.read_excel('U:\Retail Reporting\AppFolder\Retail\In\Book1.xlsx')
# domestic = domestic.ix[20:]
# domestic.columns = ['Store_Name','Vendor_Code','Season_Code','Sub_Dept','Product_Name','Color_Desc','Size_Code','Orig_Retail_Price','mtdunits_15','mtd$_15','ytdunits_15','ytd$_15','ltdunits_16','ltd$_16','ltdunits_15','ltd$_15','rcptunits_16','rcpt$_16','rcptunits_15','rcpt$_15']
# domestic = domestic.reset_index(drop=True)

print 'Harrods...'

harrods = pd.io.excel.read_excel(oldData, 0)

harrods = harrods[['DEPT','STORE','VC','SEASON','DESC','COLOR','SIZE','RETAIL','YTDU','YTD$','LTDU','LTD$']]
harrods.columns = ['Sub_Dept','Store_Name','Vendor_Code','Season_Code','Product_Name','Color_Desc','Size_Code','Orig_Retail_Price',
                   'ytdunits_16','ytd$_16','ltdunits_16','ltd$_16']
harrods['rcptunits_16'] = harrods['ltdunits_16']
harrods['rcpt$_16'] = harrods.ltdunits_16.astype(float) * harrods.Orig_Retail_Price.astype(float)

print 'Raymark sales...'

raySales = pd.read_csv('U:\Retail Reporting\AppFolder\Retail\In\\Sales Summary By Product Report.csv', 
    usecols=['Vendor_Code','Store_Name','Transaction Date','Qty_Sold_Net','Net_Sold_Retail','Product_Name','Color_Desc','Size_Code','Season_code','Sub_Dept','Orig Retail Price'], dtype={'Qty_Sold_Net':np.float64, 'Net_Sold_Retail':np.float64, 'Orig Retail Price':np.float64})
raySales['Season_Code'] = raySales.Season_code
raySales['Orig_Retail_Price'] = raySales['Orig Retail Price']
raySales['Transaction_Date'] = pd.to_datetime(raySales['Transaction Date'], format='%Y%m%d')
raySales['Month'] = raySales.Transaction_Date.dt.month
raySales['Year'] = raySales.Transaction_Date.dt.year
raySales['Day'] = raySales.Transaction_Date.dt.day

month = datetime.today().month 
thisyear = datetime.today().year
lastyear = thisyear - 1
thisday = datetime.today().day

raySales = raySales[(raySales.Transaction_Date >= date(2016, 4, 1)) | (raySales.Transaction_Date < date(lastyear, month, thisday))]

raySales['mtdunits_16'] = raySales.Qty_Sold_Net[(raySales.Month == month) & (raySales.Year == thisyear)]
raySales['mtd$_16'] = raySales.Net_Sold_Retail[(raySales.Month == month) & (raySales.Year == thisyear)]
raySales['ytdunits_16'] = raySales.Qty_Sold_Net[(raySales.Month >= 4) & (raySales.Year == thisyear)]
raySales['ytd$_16'] = raySales.Net_Sold_Retail[(raySales.Month >= 4) & (raySales.Year == thisyear)]
raySales['ltdunits_16'] = raySales.Qty_Sold_Net[(raySales.Month >= 4) & (raySales.Year == thisyear)]
raySales['ltd$_16'] = raySales.Net_Sold_Retail[(raySales.Month >= 4) & (raySales.Year == thisyear)]

raySales['mtdunits_15'] = raySales.Qty_Sold_Net[(raySales.Month == month) & (raySales.Year == lastyear) & (raySales.Day < thisday)]
raySales['mtd$_15'] = raySales.Net_Sold_Retail[(raySales.Month == month) & (raySales.Year == lastyear) & (raySales.Day < thisday)]
raySales['ytdunits_15'] = raySales.Qty_Sold_Net[(raySales.Transaction_Date < date(lastyear, month, thisday))]
raySales['ytd$_15'] = raySales.Net_Sold_Retail[(raySales.Transaction_Date < date(lastyear, month, thisday))]
raySales['ltdunits_15'] = raySales.Qty_Sold_Net[(raySales.Transaction_Date < date(lastyear, month, thisday))]
raySales['ltd$_15'] = raySales.Net_Sold_Retail[(raySales.Transaction_Date < date(lastyear, month, thisday))]

raySales = raySales.reset_index(drop=True)

print 'Raymark inventory...'
rayInv = pd.read_csv('U:\Retail Reporting\AppFolder\Retail\In\\Current On Hand Report.csv', 
    usecols=['Store_Name','Product_Name','Season_Code','Color_Desc','Size_Code','Vendor_Code','Sub_Department','Onhand_Qty','Orig_Retail_Price','Class'],
    dtype={'Onhand_Qty':np.float64, 'Orig_Retail_Price':np.float64})

rayInv.Season_Code = rayInv.Season_Code.str.strip()
rayInv.Onhand_Qty.replace(to_replace=0, value=np.nan, inplace=True)
rayInv = rayInv.dropna(subset=['Onhand_Qty'])

rayInv['Sub_Dept'] = rayInv.Sub_Department

rayInv['Onhand_Retail'] = rayInv.Onhand_Qty * rayInv.Orig_Retail_Price
rayInv.fillna(value=0, inplace=True)

rayInv = rayInv.groupby([rayInv.Store_Name, rayInv.Product_Name, rayInv.Season_Code, rayInv.Color_Desc, rayInv.Size_Code, rayInv.Vendor_Code,
    rayInv.Sub_Dept, rayInv.Orig_Retail_Price, rayInv.Class]).sum()
rayInv = rayInv[['Onhand_Qty','Onhand_Retail']]
rayInv = rayInv.reset_index()

print 'Special data...'
YTDRP = pd.io.excel.read_excel(oldData, 1)

YTDRP.columns = ['Store_Name','Vendor_Code','Season_Code','Sub_Dept','Product_Name','Color_Desc','Size_Code','Orig_Retail_Price','ytdunits_16','ytd$_16']

YTDRP['ltdunits_16'] = YTDRP.ytdunits_16
YTDRP['ltd$_16'] = YTDRP['ytd$_16']

YTDRP = YTDRP.reset_index(drop=True)

YTDRay = pd.io.excel.read_excel(oldData, 2)

YTDRay['Season_Code'] = YTDRay.Season_code
YTDRay['Orig_Retail_Price'] = YTDRay['Orig Retail Price']

YTDRay['ytdunits_16'] = YTDRay['Qty_Sold_Net']
YTDRay['ytd$_16'] = YTDRay['Net_Sold_Retail']
YTDRay['ltdunits_16'] = YTDRay['Qty_Sold_Net']
YTDRay['ltd$_16'] = YTDRay['Net_Sold_Retail']
YTDRay['Month'] = 3

YTDRay = YTDRay.reset_index(drop=True)

SalesDB = pd.merge(harrods, pd.merge(raySales, pd.merge(rayInv, pd.merge(YTDRP, YTDRay, how='outer'), how='outer'), how='outer'), how='outer')

SalesDB['Class_Num'] = np.nan
SalesDB['Fabric'] = np.nan
SalesDB['rcptunits_16'] = 0
SalesDB['rcpt$_16'] = 0
SalesDB['rcptunits_15'] = 0
SalesDB['rcpt$_15'] = 0

SalesDB[['Orig_Retail_Price','Onhand_Qty','mtdunits_16','mtd$_16','mtdunits_15','mtd$_15','ytdunits_16','ytd$_16','ytdunits_15','ytd$_15','ltdunits_16','ltd$_16','ltdunits_15','ltd$_15','rcptunits_15','rcpt$_15','rcptunits_16','rcpt$_16','Onhand_Qty','Onhand_Retail']] = SalesDB[['Orig_Retail_Price','Onhand_Qty','mtdunits_16','mtd$_16','mtdunits_15','mtd$_15','ytdunits_16','ytd$_16','ytdunits_15','ytd$_15','ltdunits_16','ltd$_16','ltdunits_15','ltd$_15','rcptunits_15','rcpt$_15','rcptunits_16','rcpt$_16','Onhand_Qty','Onhand_Retail']].fillna(value=0)

SalesDB.rcptunits_16 = SalesDB.Onhand_Qty + SalesDB.ltdunits_16
SalesDB['rcpt$_16'] = SalesDB.rcptunits_16 * SalesDB.Orig_Retail_Price

SalesDB = SalesDB[['Store_Name','Vendor_Code','Season_Code','Sub_Dept','Product_Name','Color_Desc','Size_Code','Orig_Retail_Price','mtdunits_16','mtd$_16','mtdunits_15','mtd$_15','ytdunits_16','ytd$_16','ytdunits_15','ytd$_15','ltdunits_16','ltd$_16','ltdunits_15','ltd$_15','rcptunits_15','rcpt$_15','rcptunits_16','rcpt$_16','Onhand_Qty','Onhand_Retail','Class_Num','Fabric','Transaction_Date','Month']]


print 'Cleaning data...'
SalesDB = processData(SalesDB)

# markdown = pd.io.excel.read_excel('U:\Retail Reporting\AppFolder\Retail\In\\Sales Markdown Detail Report.xls')
# markdownly = pd.io.excel.read_excel('U:\Retail Reporting\AppFolder\Retail\In\\Sales Markdown Detail Report ly.xls')
# markdown = pd.merge(markdown, markdownly, how='outer')
# markdown.Store_Name.replace(to_replace="CHNY Madison Ave.", value="MAD", inplace=True)
# markdown.Store_Name.replace(to_replace="CHNY Rodeo Drive", value="LA", inplace=True)
# markdown.Store_Name.replace(to_replace="CHNY Melrose Pl.", value="LA", inplace=True)
# markdown.Store_Name.replace(to_replace="CHNY Dallas", value="DAL", inplace=True)
# markdown.Store_Name.replace(to_replace="CHNY Harrods", value="HAR", inplace=True)
# markdown.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\markdown.csv')

if EOMREPORT:
    print 'Style selling...'
    sales = SalesDB

    wholesaleStyles = pd.read_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\StyleSalesDB.csv', usecols=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18])

    wholesaleStyles.columns = ['index','Season_Code','Product_Name','Color_Desc','Orig_Retail_Price','wtdunits','wtdsales','Qty_Sold_Net','Net_Sold_Retail','rcpts','rcptsales','Onhand_Qty','Onhand_Retail','oounits','ooretail','Store_Name','lastrcvd','classnum','Fabric']
    wholesaleStyles = wholesaleStyles[['Season_Code','Product_Name','Orig_Retail_Price','Qty_Sold_Net','Net_Sold_Retail','Onhand_Qty','Onhand_Retail','Store_Name']]

    sales = sales[['Store_Name','ytdunits_16','ytd$_16','ltdunits_16','ltd$_16','Product_Name','Color_Desc','Season_Code','Sub_Dept','Orig_Retail_Price','Vendor_Code','Onhand_Retail','Onhand_Qty']]
    sales.columns = ['Store_Name','Qty_Sold_Net','Net_Sold_Retail','Qty_Sold_Netltd','Net_Sold_Retailltd','Product_Name','Color_Desc','Season_Code','Sub_Dept','Orig_Retail_Price','Vendor_Code','Onhand_Retail','Onhand_Qty']

    wholesaleStyles.Net_Sold_Retail = wholesaleStyles.Net_Sold_Retail * 1000
    wholesaleStyles.Onhand_Retail = wholesaleStyles.Onhand_Retail * 1000

    wholesaleStyles.Season_Code = wholesaleStyles.Product_Name.map(seasonDict)

    i = 0
    for name in wholesaleStyles.Product_Name:
        if (name in ['201P','208P','209P','204P','2057ZNC','2007CDC','2017CDC','2160CJM','2027GDK','2017GDK','2037GDK','2047LAC','2040ATAEMB','2010PTS','2140TEM','003P','001P','002P','0010SWTEMB','1007LEA','1027LEA','1017LAC','1010PTS','1040PTS','3037LBA','3009RAO','3027DNM','3027COT','3037LEA','3027LEA','3057JTX','3007LAC','3020PTS','3040PTS','3030PTS','3050CJT','3080MAU','4003LBA','4013LBA','4023LBA','4007COT','4030CPY','4040CPY','5087DNM','5097DNM','5009RAY','5039RAO','5029RAO','5037COT','5047GDK','5057COT','5057SLK','5080MAU','5150MAU','5017LAC','5007LAC','5027LAC','5160FIN','5140PTSEMB','7040CJT','7050CCC','7010CJA','7110MAU','7070RCC','7020PTS','7030PTSEMB','7120PTS','9017BLT','9017NG','9017GDH','9007BLT']):

            wholesaleStyles.loc[i, 'Season_Code'] = 'ICON'

            i = i + 1

    wholesaleStyles['Vendor_Code'] = 'CH'

    sellThrus = pd.merge(wholesaleStyles, pd.merge(sales, processBookings(), how='outer'), how='outer')
    sellThrus = sellThrus[(sellThrus.Store_Name != 'In Transit virtual Store_Name')]
    sellThrus = sellThrus.reset_index(drop=True)
    sellThrus = cleanColors(sellThrus)

    sellThrus['Ship_Date'] = ''
    sellThrus['Class'] = ''
    sellThrus.Ship_Date = sellThrus.Product_Name.map(dateDict)
    sellThrus.Orig_Retail_Price = sellThrus.Product_Name.map(priceDict)
    sellThrus.Sub_Dept = sellThrus.Product_Name.map(descDict)
    sellThrus.Class = sellThrus.Product_Name.map(classDict)
    sellThrus['Qty_Ord_Retail'] = sellThrus.Qty_Ordered * sellThrus.Orig_Retail_Price

    sellThrus = sellThrus[['Season_Code','Product_Name','Class','Sub_Dept','Color_Desc','Orig_Retail_Price','Ship_Date', \
        'Store_Name','Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Sold_Netltd','Net_Sold_Retailltd','Qty_Ordered','Qty_Ord_Retail','Vendor_Code']]

    sellThrus.Net_Sold_Retail = pd.to_numeric(sellThrus.Net_Sold_Retail, errors='coerce')

    sellThrus.Season_Code.replace(to_replace="PFP", value="16P", inplace=True)
    sellThrus.fillna(value=0, inplace=True)

    bridalUniqueStyles = sellThrus[(sellThrus.Vendor_Code.isin(['BRIDAL','BEHR','EH','OSSAI'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDBRD','nan','SOBRIDAL','UPCHARGEBRIDAL','UPCHARGEBRIDAL2','UPCHARGEBRIDAL3']))]
    bridalUniqueStyles = bridalUniqueStyles.groupby([bridalUniqueStyles.Season_Code, bridalUniqueStyles.Product_Name, bridalUniqueStyles.Class, \
            bridalUniqueStyles.Sub_Dept, bridalUniqueStyles.Color_Desc, bridalUniqueStyles.Orig_Retail_Price]).sum()
    bridalUniqueStyles = bridalUniqueStyles[['Onhand_Retail','Onhand_Qty','Net_Sold_Retailltd','Qty_Sold_Netltd','Qty_Ordered','Qty_Ord_Retail']]
    bridalUniqueStyles.replace(to_replace="0", value=np.nan, inplace=True)
    bridalUniqueStyles.replace(to_replace=0, value=np.nan, inplace=True)
    bridalUniqueStyles.dropna(how='all', inplace=True)

    accessoriesUniqueStyles = sellThrus[~sellThrus.Vendor_Code.isin(['BRIDAL','BEHR','EH','OSSAI','CHNY','CH',np.nan])]
    accessoriesUniqueStyles = accessoriesUniqueStyles.groupby([accessoriesUniqueStyles.Season_Code, accessoriesUniqueStyles.Product_Name, accessoriesUniqueStyles.Class, \
            accessoriesUniqueStyles.Sub_Dept, accessoriesUniqueStyles.Color_Desc, accessoriesUniqueStyles.Orig_Retail_Price]).sum()
    accessoriesUniqueStyles = accessoriesUniqueStyles[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    accessoriesUniqueStyles.replace(to_replace="0", value=np.nan, inplace=True)
    accessoriesUniqueStyles.replace(to_replace=0, value=np.nan, inplace=True)
    accessoriesUniqueStyles.dropna(subset=['Onhand_Qty'], inplace=True)
    accessoriesUniqueStyles.dropna(how='all', inplace=True)

    IconUniqueStyles = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['ICON'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW'])) & (~sellThrus.Store_Name.isin(['bg','sfa','nm']))]
    IconUniqueStyles = IconUniqueStyles.groupby([IconUniqueStyles.Season_Code, IconUniqueStyles.Product_Name, IconUniqueStyles.Class, \
            IconUniqueStyles.Sub_Dept, IconUniqueStyles.Color_Desc, IconUniqueStyles.Orig_Retail_Price]).sum()
    IconUniqueStyles = IconUniqueStyles[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    IconUniqueStyles.replace(to_replace="0", value=np.nan, inplace=True)
    IconUniqueStyles.replace(to_replace=0, value=np.nan, inplace=True)
    IconUniqueStyles.dropna(how='all', inplace=True)

    PF16UniqueStyles = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['16P','PFP'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW']))]
    PF16UniqueStyles = PF16UniqueStyles.groupby([PF16UniqueStyles.Season_Code, PF16UniqueStyles.Product_Name, PF16UniqueStyles.Class, \
            PF16UniqueStyles.Sub_Dept, PF16UniqueStyles.Color_Desc, PF16UniqueStyles.Orig_Retail_Price]).sum()
    PF16UniqueStyles = PF16UniqueStyles[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    PF16UniqueStyles.replace(to_replace="0", value=np.nan, inplace=True)
    PF16UniqueStyles.replace(to_replace=0, value=np.nan, inplace=True)
    PF16UniqueStyles.dropna(how='all', inplace=True)

    FA16UniqueStyles = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['16F'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW']))]
    FA16UniqueStyles = FA16UniqueStyles.groupby([FA16UniqueStyles.Season_Code, FA16UniqueStyles.Product_Name, FA16UniqueStyles.Class, \
            FA16UniqueStyles.Sub_Dept, FA16UniqueStyles.Color_Desc, FA16UniqueStyles.Orig_Retail_Price]).sum()
    FA16UniqueStyles = FA16UniqueStyles[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    FA16UniqueStyles.replace(to_replace="0", value=np.nan, inplace=True)
    FA16UniqueStyles.replace(to_replace=0, value=np.nan, inplace=True)
    FA16UniqueStyles.dropna(how='all', inplace=True)

    bridalSellThrus = sellThrus[(sellThrus.Vendor_Code.isin(['BRIDAL','BEHR','EH','OSSAI'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDBRD','nan','SOBRIDAL','UPCHARGEBRIDAL','UPCHARGEBRIDAL2','UPCHARGEBRIDAL3']))]
    bridalSellThrus = bridalSellThrus.groupby([bridalSellThrus.Season_Code, bridalSellThrus.Product_Name, bridalSellThrus.Class, \
            bridalSellThrus.Sub_Dept, bridalSellThrus.Color_Desc, bridalSellThrus.Orig_Retail_Price, bridalSellThrus.Ship_Date, bridalSellThrus.Store_Name]).sum()
    bridalSellThrus = bridalSellThrus[['Onhand_Retail','Onhand_Qty','Net_Sold_Retailltd','Qty_Sold_Netltd','Qty_Ordered','Qty_Ord_Retail']]
    bridalSellThrus.replace(to_replace="0", value=np.nan, inplace=True)
    bridalSellThrus.replace(to_replace=0, value=np.nan, inplace=True)
    bridalSellThrus.dropna(how='all', inplace=True)

    accessoriesSellThrus = sellThrus[~sellThrus.Vendor_Code.isin(['BRIDAL','BEHR','EH','OSSAI','CHNY','CH',np.nan])]
    accessoriesSellThrus = accessoriesSellThrus.groupby([accessoriesSellThrus.Season_Code, accessoriesSellThrus.Product_Name, accessoriesSellThrus.Class, \
            accessoriesSellThrus.Sub_Dept, accessoriesSellThrus.Color_Desc, accessoriesSellThrus.Orig_Retail_Price, accessoriesSellThrus.Ship_Date, accessoriesSellThrus.Store_Name]).sum()
    accessoriesSellThrus = accessoriesSellThrus[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    accessoriesSellThrus.replace(to_replace="0", value=np.nan, inplace=True)
    accessoriesSellThrus.replace(to_replace=0, value=np.nan, inplace=True)
    accessoriesUniqueStyles.dropna(subset=['Onhand_Qty'], inplace=True)
    accessoriesSellThrus.dropna(how='all', inplace=True)

    IconSellThrus = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['ICON'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW'])) & (~sellThrus.Store_Name.isin(['bg','sfa','nm']))]
    IconSellThrus = IconSellThrus.groupby([IconSellThrus.Season_Code, IconSellThrus.Product_Name, IconSellThrus.Class, \
            IconSellThrus.Sub_Dept, IconSellThrus.Color_Desc, IconSellThrus.Orig_Retail_Price, IconSellThrus.Ship_Date, IconSellThrus.Store_Name]).sum()
    IconSellThrus = IconSellThrus[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    IconSellThrus.replace(to_replace="0", value=np.nan, inplace=True)
    IconSellThrus.replace(to_replace=0, value=np.nan, inplace=True)
    IconSellThrus.dropna(how='all', inplace=True)

    PF16SellThrus = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['16P','PFP'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW']))]
    PF16SellThrus = PF16SellThrus.groupby([PF16SellThrus.Season_Code, PF16SellThrus.Product_Name, PF16SellThrus.Class, \
            PF16SellThrus.Sub_Dept, PF16SellThrus.Color_Desc, PF16SellThrus.Orig_Retail_Price, PF16SellThrus.Ship_Date, PF16SellThrus.Store_Name]).sum()
    PF16SellThrus = PF16SellThrus[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    PF16SellThrus.replace(to_replace="0", value=np.nan, inplace=True)
    PF16SellThrus.replace(to_replace=0, value=np.nan, inplace=True)
    PF16SellThrus.dropna(how='all', inplace=True)

    FA16SellThrus = sellThrus[(sellThrus.Vendor_Code.isin(['CHNY','CH',np.nan])) & (sellThrus['Season_Code'].isin(['16F'])) & (~sellThrus.Product_Name.isin(['BRDUPCHARGE','INVALIDRTW','nan','SORTW']))]
    FA16SellThrus = FA16SellThrus.groupby([FA16SellThrus.Season_Code, FA16SellThrus.Product_Name, FA16SellThrus.Class, \
            FA16SellThrus.Sub_Dept, FA16SellThrus.Color_Desc, FA16SellThrus.Orig_Retail_Price, FA16SellThrus.Ship_Date, FA16SellThrus.Store_Name]).sum()
    FA16SellThrus = FA16SellThrus[['Onhand_Retail','Onhand_Qty','Net_Sold_Retail','Qty_Sold_Net','Qty_Ordered','Qty_Ord_Retail']]
    FA16SellThrus.replace(to_replace="0", value=np.nan, inplace=True)
    FA16SellThrus.replace(to_replace=0, value=np.nan, inplace=True)
    FA16SellThrus.dropna(how='all', inplace=True)

    bridalUniqueStyles.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\bridalUniqueStyles.csv')
    accessoriesUniqueStyles.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\accessoriesUniqueStyles.csv')
    IconUniqueStyles.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\IconUniqueStyles.csv')
    PF16UniqueStyles.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\PF16UniqueStyles.csv')
    FA16UniqueStyles.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\FA16UniqueStyles.csv')

    bridalSellThrus.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\bridalSellThrus.csv')
    accessoriesSellThrus.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\accessoriesSellThrus.csv')
    IconSellThrus.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\IconSellThrus.csv')
    PF16SellThrus.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\PF16SellThrus.csv')
    FA16SellThrus.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\FA16SellThrus.csv')

    print 'Icon selling...'
    analyzeIconSales(SalesDB).to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\IconReplenishmentData.csv')

tempdf = pd.read_csv('U:\Retail Reporting\AppFolder\Retail\In\\orliweb.txt', encoding='utf-16', sep='\t')
i = 0
for name in tempdf.Style:
    if name in ['3027GDL','3027GQI','3037GDL','3037GQI','3973GQJ','4003GQI','4003GDL','4063GDL','5027GDL','5037GDL','5077GDL','5483GQI','5633GDL','5633GQI','5793GQJ','5803GQJ']:
        tempdf.loc[i, 'Season'] = 'PFP'

    elif (name in ['201P','208P','209P','204P','2057ZNC','2007CDC','2017CDC','2160CJM','2027GDK','2017GDK','2037GDK','2047LAC','2040ATAEMB','2010PTS','2140TEM','003P','001P','002P','0010SWTEMB','1007LEA','1027LEA','1017LAC','1010PTS','1040PTS','3037LBA','3009RAO','3027DNM','3027COT','3037LEA','3027LEA','3057JTX','3007LAC','3020PTS','3040PTS','3030PTS','3050CJT','3080MAU','4003LBA','4013LBA','4023LBA','4007COT','4030CPY','4040CPY','5087DNM','5097DNM','5009RAY','5039RAO','5029RAO','5037COT','5047GDK','5057COT','5057SLK','5080MAU','5150MAU','5017LAC','5007LAC','5027LAC','5160FIN','5140PTSEMB','7040CJT','7050CCC','7010CJA','7110MAU','7070RCC','7020PTS','7030PTSEMB','7120PTS','9017BLT','9017NG','9017GDH','9007BLT']) & (tempdf.loc[i, 'Season'] in ['CPC','DAC','ESC','EVC']):
            tempdf.loc[i, 'Season'] = 'ICON'

    i = i + 1
tempdf.to_csv('U:\Retail Reporting\AppFolder\Retail\In\\orliweb.csv')


print "Updating the output files!"

bridal = SalesDB[(SalesDB.Vendor_Code.isin(['BRIDAL','EH','BEHR','OSSAI']))]
bridal['filler'] = np.nan
bridal['Year'] = 2016
bridal = bridal[['Store_Name','ytd$_16','filler','Year','Month']]
bridal = bridal.dropna(subset=['ytd$_16','Month'])
bridal.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\bridal.csv', index=False)

bridalUnits = SalesDB[(SalesDB.Vendor_Code.isin(['BRIDAL','EH','BEHR','OSSAI']))]
bridalUnits['filler_1'] = np.nan
bridalUnits['Year'] = 2016
bridalUnits = bridalUnits[['Vendor_Code','Store_Name','Year','filler_1','ytdunits_16','Product_Name','Season_Code','Sub_Dept','Month']]
bridalUnits = bridalUnits.dropna(subset=['ytdunits_16','Month'])
bridalUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\bridalUnits.csv')

# fabrics = SalesDB[['Season_Code','Fabric','mtdunits_16']]
# fabrics = fabrics[(fabrics.Season_Code == '16S') | (fabrics.Season_Code == '16R') | (fabrics.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')]
# fabrics = fabrics.groupby([fabrics.Fabric])['mtdunits_16'].sum().sort_values(ascending=False)[:9]
# fabrics.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\fabrics.csv')

# colors = SalesDB[['Season_Code','Color_Desc','mtdunits_16']]
# colors = colors[(colors.Season_Code == '16S') | (colors.Season_Code == '16R') | (colors.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')]
# colors = colors.groupby([colors['Color_Desc']])['mtdunits_16'].sum().sort_values(ascending=False)[:9]
# colors.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\colors.csv')

Size_Codes = SalesDB[['Vendor_Code','Size_Code','mtdunits_16']][(SalesDB.Vendor_Code == 'CH')]
Size_Codes.Size_Code = Size_Codes.Size_Code.apply(str)
Size_Codes.mtdunits_16 = Size_Codes.mtdunits_16.apply(float)
totalUnits = Size_Codes.mtdunits_16.sum()
Size_Codes.mtdunits_16 = Size_Codes.mtdunits_16 / totalUnits
Size_Codes = Size_Codes.groupby([Size_Codes['Size_Code']])['mtdunits_16'].sum().sort_values(ascending=False)[:10]
Size_Codes.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\allsizes.csv')

topPrefallUnits = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')]
topPrefallUnits.mtdunits_16 = pd.to_numeric(topPrefallUnits.mtdunits_16, errors='coerce')
topPrefallUnits['mtd$_16'] = pd.to_numeric(topPrefallUnits['mtd$_16'], errors='coerce')
topPrefallUnits = topPrefallUnits.groupby([topPrefallUnits['Product_Name']]).sum()
topPrefallUnits = topPrefallUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
topPrefallUnits = topPrefallUnits.reset_index()
topPrefallUnits = topPrefallUnits.replace(to_replace=0, value=np.nan).dropna()
i = 0
topPrefallUnits['Class_Indicator'] = np.nan
for style in topPrefallUnits.Product_Name.values:
    topPrefallUnits.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topPrefallUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topPrefallUnits.csv', index=False)

topPrefallSales = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')]
topPrefallSales.mtdunits_16 = pd.to_numeric(topPrefallSales.mtdunits_16, errors='coerce')
topPrefallSales['mtd$_16'] = pd.to_numeric(topPrefallSales['mtd$_16'], errors='coerce')
topPrefallSales = topPrefallSales.groupby([topPrefallSales['Product_Name']]).sum()
topPrefallSales = topPrefallSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
topPrefallSales = topPrefallSales.reset_index()
topPrefallSales = topPrefallSales.replace(to_replace=0, value=np.nan).dropna()
i = 0
topPrefallSales['Class_Indicator'] = np.nan
for style in topPrefallSales.Product_Name.values:
    topPrefallSales.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topPrefallSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topPrefallSales.csv', index=False)

topFallUnits = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16F')]
topFallUnits.mtdunits_16 = pd.to_numeric(topFallUnits.mtdunits_16, errors='coerce')
topFallUnits['mtd$_16'] = pd.to_numeric(topFallUnits['mtd$_16'], errors='coerce')
topFallUnits = topFallUnits.groupby([topFallUnits['Product_Name']]).sum()
topFallUnits = topFallUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
topFallUnits = topFallUnits.reset_index()
topFallUnits = topFallUnits.replace(to_replace=0, value=np.nan).dropna()
i = 0
topFallUnits['Class_Indicator'] = np.nan
for style in topFallUnits.Product_Name.values:
    topFallUnits.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topFallUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topFallUnits.csv', index=False)

topFallSales = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16F')]
topFallSales.mtdunits_16 = pd.to_numeric(topFallSales.mtdunits_16, errors='coerce')
topFallSales['mtd$_16'] = pd.to_numeric(topFallSales['mtd$_16'], errors='coerce')
topFallSales = topFallSales.groupby([topFallSales['Product_Name']]).sum()
topFallSales = topFallSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
topFallSales = topFallSales.reset_index()
topFallSales = topFallSales.replace(to_replace=0, value=np.nan).dropna()
i = 0
topFallSales['Class_Indicator'] = np.nan
for style in topFallSales.Product_Name.values:
    topFallSales.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topFallSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topFallSales.csv', index=False)


topIconUnits = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == 'ICON') & (SalesDB.Vendor_Code == 'CH')]
topIconUnits.mtdunits_16 = pd.to_numeric(topIconUnits.mtdunits_16, errors='coerce')
topIconUnits['mtd$_16'] = pd.to_numeric(topIconUnits['mtd$_16'], errors='coerce')
topIconUnits = topIconUnits.groupby([topIconUnits['Product_Name']]).sum()
topIconUnits = topIconUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
topIconUnits = topIconUnits.reset_index()
topIconUnits = topIconUnits.replace(to_replace=0, value=np.nan).dropna()
i = 0
topIconUnits['Class_Indicator'] = np.nan
for style in topIconUnits.Product_Name.values:
    topIconUnits.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topIconUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topIconUnits.csv', index=False)

topIconSales = SalesDB[['Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == 'ICON') & (SalesDB.Vendor_Code == 'CH')]
topIconSales.mtdunits_16 = pd.to_numeric(topIconSales.mtdunits_16, errors='coerce')
topIconSales['mtd$_16'] = pd.to_numeric(topIconSales['mtd$_16'], errors='coerce')
topIconSales = topIconSales.groupby([topIconSales['Product_Name']]).sum()
topIconSales = topIconSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
topIconSales = topIconSales.reset_index()
topIconSales = topIconSales.replace(to_replace=0, value=np.nan).dropna()
i = 0
topIconSales['Class_Indicator'] = np.nan
for style in topIconSales.Product_Name.values:
    topIconSales.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
topIconSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topIconSales.csv', index=False)

stores = ['MAD','DAL','LA','HAR']
for store in stores:

    topPrefallUnits = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][((SalesDB.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')) & (SalesDB.Store_Name == store)]
    topPrefallUnits.mtdunits_16 = pd.to_numeric(topPrefallUnits.mtdunits_16, errors='coerce')
    topPrefallUnits['mtd$_16'] = pd.to_numeric(topPrefallUnits['mtd$_16'], errors='coerce')
    topPrefallUnits = topPrefallUnits.groupby([topPrefallUnits['Product_Name']]).sum()
    topPrefallUnits = topPrefallUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
    topPrefallUnits = topPrefallUnits.reset_index()
    topPrefallUnits = topPrefallUnits.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topPrefallUnits['Class_Indicator'] = np.nan
    for style in topPrefallUnits.Product_Name.values:
        topPrefallUnits.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topPrefallUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topPrefallUnits'+store+'.csv', index=False)

    topPrefallSales = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][((SalesDB.Season_Code == '16P') | (SalesDB.Season_Code == 'PFP')) & (SalesDB.Store_Name == store)]
    topPrefallSales.mtdunits_16 = pd.to_numeric(topPrefallSales.mtdunits_16, errors='coerce')
    topPrefallSales['mtd$_16'] = pd.to_numeric(topPrefallSales['mtd$_16'], errors='coerce')
    topPrefallSales = topPrefallSales.groupby([topPrefallSales['Product_Name']]).sum()
    topPrefallSales = topPrefallSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
    topPrefallSales = topPrefallSales.reset_index()
    topPrefallSales = topPrefallSales.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topPrefallSales['Class_Indicator'] = np.nan
    for style in topPrefallSales.Product_Name.values:
        topPrefallSales.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topPrefallSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topPrefallSales'+store+'.csv', index=False)

    topFallUnits = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16F') & (SalesDB.Store_Name == store)]
    topFallUnits.mtdunits_16 = pd.to_numeric(topFallUnits.mtdunits_16, errors='coerce')
    topFallUnits['mtd$_16'] = pd.to_numeric(topFallUnits['mtd$_16'], errors='coerce')
    topFallUnits = topFallUnits.groupby([topFallUnits['Product_Name']]).sum()
    topFallUnits = topFallUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
    topFallUnits = topFallUnits.reset_index()
    topFallUnits = topFallUnits.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topFallUnits['Class_Indicator'] = np.nan
    for style in topFallUnits.Product_Name.values:
        topFallUnits.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topFallUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topFallUnits'+store+'.csv', index=False)

    topFallSales = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == '16F') & (SalesDB.Store_Name == store)]
    topFallSales.mtdunits_16 = pd.to_numeric(topFallSales.mtdunits_16, errors='coerce')
    topFallSales['mtd$_16'] = pd.to_numeric(topFallSales['mtd$_16'], errors='coerce')
    topFallSales = topFallSales.groupby([topFallSales['Product_Name']]).sum()
    topFallSales = topFallSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
    topFallSales = topFallSales.reset_index()
    topFallSales = topFallSales.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topFallSales['Class_Indicator'] = np.nan
    for style in topFallSales.Product_Name.values:
        topFallSales.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topFallSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topFallSales'+store+'.csv', index=False)

    topIconUnits = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == 'ICON') & (SalesDB.Vendor_Code == 'CH') & (SalesDB.Store_Name == store)]
    topIconUnits.mtdunits_16 = pd.to_numeric(topIconUnits.mtdunits_16, errors='coerce')
    topIconUnits['mtd$_16'] = pd.to_numeric(topIconUnits['mtd$_16'], errors='coerce')
    topIconUnits = topIconUnits.groupby([topIconUnits['Product_Name']]).sum()
    topIconUnits = topIconUnits.sort_values(by=['mtdunits_16','mtd$_16'], ascending=False)[:5]
    topIconUnits = topIconUnits.reset_index()
    topIconUnits = topIconUnits.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topIconUnits['Class_Indicator'] = np.nan
    for style in topIconUnits.Product_Name.values:
        topIconUnits.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topIconUnits.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topIconUnits'+store+'.csv', index=False)

    topIconSales = SalesDB[['Store_Name','Season_Code','Product_Name','mtdunits_16','mtd$_16']][(SalesDB.Season_Code == 'ICON') & (SalesDB.Vendor_Code == 'CH') & (SalesDB.Store_Name == store)]
    topIconSales.mtdunits_16 = pd.to_numeric(topIconSales.mtdunits_16, errors='coerce')
    topIconSales['mtd$_16'] = pd.to_numeric(topIconSales['mtd$_16'], errors='coerce')
    topIconSales = topIconSales.groupby([topIconSales['Product_Name']]).sum()
    topIconSales = topIconSales.sort_values(by=['mtd$_16','mtdunits_16'], ascending=False)[:5]
    topIconSales = topIconSales.reset_index()
    topIconSales = topIconSales.replace(to_replace=0, value=np.nan).dropna()
    i = 0
    topIconSales['Class_Indicator'] = np.nan
    for style in topIconSales.Product_Name.values:
        topIconSales.loc[i, 'Class_Indicator'] = style[0:1]
        i = i + 1
    topIconSales.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\topIconSales'+store+'.csv', index=False)

    topStoreSizes = SalesDB[['Store_Name','Vendor_Code','Size_Code','mtdunits_16']][(SalesDB.Vendor_Code == 'CH') & (SalesDB.Store_Name == store)]
    topStoreSizes.Size_Code = topStoreSizes.Size_Code.apply(str)
    topStoreSizes.mtdunits_16 = topStoreSizes.mtdunits_16.apply(float)
    totalUnits = topStoreSizes.mtdunits_16.sum()
    topStoreSizes.mtdunits_16 = topStoreSizes.mtdunits_16 / totalUnits
    topStoreSizes = topStoreSizes.groupby([topStoreSizes['Size_Code']])['mtdunits_16'].sum().sort_values(ascending=False)[:10]
    topStoreSizes.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\\'+store+'sizes.csv')


clients.Qty_Sold = pd.to_numeric(clients.Qty_Sold, errors='coerce')
clients.Ext_Price_Sold = pd.to_numeric(clients.Ext_Price_Sold, errors='coerce')
clients.Price_sold = pd.to_numeric(clients.Price_sold, errors='coerce')

clients['transdate'] = pd.to_datetime(clients['transdate'], format='%Y%m%d')
clients['year'] = pd.DatetimeIndex(clients['transdate']).year
clients['month'] = pd.DatetimeIndex(clients['transdate']).month
clients['day'] = pd.DatetimeIndex(clients['transdate']).day
clients['returns'] = clients.Ext_Price_Sold < 0
clients['mtdunits'] = clients.Qty_Sold[(clients.month == month)]
clients['mtd$'] = clients.Price_sold[(clients.month == month)]

clients['ytdunits'] = clients.Qty_Sold[(clients.year == thisyear)]
clients['ytd$'] = clients.Price_sold[(clients.year == thisyear)]

clients = clients[['Store_Name','VIPName','Country','transdate','year','month','day','Price_sold','Qty_Sold','returns','Prov','mtdunits','mtd$','ytdunits','ytd$']]
clients.Store_Name.replace(to_replace="CHNY Madison Ave.", value="MAD", inplace=True)
clients.Store_Name.replace(to_replace="CHNY Rodeo Drive", value="LA", inplace=True)
clients.Store_Name.replace(to_replace="CHNY Melrose Pl.", value="LA", inplace=True)
clients.Store_Name.replace(to_replace="CHNY Dallas", value="DAL", inplace=True)
clients.Store_Name.replace(to_replace="CHNY Harrods", value="HAR", inplace=True)

clients.VIPName.replace(to_replace='FARFETCH DONOTUSE', value='FARFETCH', inplace=True)
clients.VIPName.replace(to_replace='FARFETCH DONOTUSE.', value='FARFETCH', inplace=True)
clients.Country.replace(to_replace='United States of America', value='United States', inplace=True)
clients.Country.replace(to_replace='United Arab Emirates', value='UAE', inplace=True)
clients.Country.replace(to_replace='Dominican Republic', value='Dominican Rep.', inplace=True)

stateCleaner(clients[(clients.Country == 'United States')])

clients.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\clients.csv')

SalesDB = SalesDB[['Store_Name','Vendor_Code','Season_Code','Sub_Dept','Product_Name','Color_Desc','Size_Code','Orig_Retail_Price','mtdunits_16','mtd$_16',
    'mtdunits_15','mtd$_15','ytdunits_16','ytd$_16','ytdunits_15','ytd$_15','ltdunits_16','ltd$_16','ltdunits_15',
    'ltd$_15','rcptunits_15','rcpt$_15','rcptunits_16','rcpt$_16','Onhand_Qty','Onhand_Retail','Class_Num','Fabric']]

SalesDB.to_csv('U:\Retail Reporting\AppFolder\Retail\Out\YummyData\SalesDB.csv')

print "All done!"

if PRODUCTION: 
    Tkinter.Tk().withdraw()
    tkMessageBox.showinfo('It\'s done!','Retail Fox is done.')


