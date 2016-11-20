import pandas as pd
import numpy as np
import datetime as dt
import shutil
from os import listdir, chdir, mkdir
from os.path import isfile, join, isdir
from openpyxl import load_workbook
import dateutil.parser as dparser
from ColorComb import combColors
import Tkinter, tkMessageBox


pd.set_option('display.max_columns', 100)

def dealWithFiles():
    Tkinter.Tk().withdraw()
    if tkMessageBox.askyesno('End of process!', 'Should we move the files? \n\n'):

        now = dt.datetime.now()
        newdir = 'U:\Retail Reporting\AppFolder\Wholesale\Out\UsedFiles\\' + str(now.date())
        if not isdir(newdir):
          mkdir(newdir)

        for f in rawData:
            try:
                shutil.move(directory + f, newdir + '\\' + f)
            except IOError:
                continue

print "Away we go!"

directory = 'U:\Retail Reporting\AppFolder\Wholesale\In\\'
rawData = [f for f in listdir(directory) if isfile(join(directory, f))]
attrMap = 'U:\Retail Reporting\AppFolder\Retail\In\AttrMap.xlsx'

assert len(rawData) >= 8, 'Not enough wholesale files, boss!'

BGDataPF = [data for data in rawData if 'F17 DETAIL SELLING' in data][0]
NMPrefallData = [data for data in rawData if 'Store Prefall' in data][0]
NMPFPData = [data for data in rawData if 'Store Early Prefall' in data][0]
NMSpecialData = [data for data in rawData if 'Store Specials' in data][0]
NMStyleData = [data for data in rawData if 'F17 HERRERA BY' in data][0]
# NMOnlineData = [data for data in rawData if 'NMD Vendor Selling' in data][0]
SaksStoreData = [data for data in rawData if 'By Door Selling' in data][0]
SaksStyleData = [data for data in rawData if 'Report by UDA' in data][0]
# SaksSpecialData = [data for data in rawData if (' 527 ' in data) | ('SPO' in data)][0]

# raise
# SaksStoresFiscal = ''
# SaksOmniFiscal = ''
# NMFiscal = ''
# try:
#   SaksStoresFiscal = [data for data in rawData if 'JULY EOM' in data][0]
#   SaksStoresFiscal = pd.io.excel.read_excel(directory+SaksStoresFiscal, header=None)

#   SaksOmniFiscal = [data for data in rawData if 'Omni-Channel' in data][0]
#   SaksOmniFiscal = pd.io.excel.read_excel(directory+SaksOmniFiscal, header=None)

#   NMFiscal = [data for data in rawData if 'NMFiscal' in data][0]
#   NMFiscal = pd.io.excel.read_excel(directory+NMFiscal, header=None)
#   NMFiscal[0].fillna(method='ffill', inplace=True)

#   SaksStoresFiscal.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\SaksStoresFiscal.csv', index=False)
#   SaksOmniFiscal.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\SaksOmniFiscal.csv', index=False)
#   NMFiscal.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\NMFiscal.csv', index=False)

# except IndexError:
#   pass

BGStorePF = pd.io.excel.read_excel(directory+BGDataPF, 1)
assert (BGStorePF.loc[175][0] == 'PRE-FALL 17 EARLY'), 'Something\'s off with Prefall, boss!'
assert (BGStorePF.loc[359][0] == 'PRE-FALL 17 STOCK'), 'Something\'s off with Prefall Stock, boss!'
assert (BGStorePF.loc[386][0] == 'PRE-FALL 17 CONSIGNMENT'), 'Something\'s off with Prefall Consignment, boss!'

BGStorePF = BGStorePF.drop(BGStorePF.columns[[0,2,3,4,6,7,8,9,10,11,12,13,14,15,16,17,22,23,24,25,26,27,28,29]],axis=1)

BGStorePF.columns = ['Group_Code','Product_Name','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD']
BGStorePF.Product_Name.replace(to_replace='VENDOR STYLE', value=np.nan, inplace=True)

BGStorePF = BGStorePF.dropna(subset=['Product_Name'])
BGStorePF = BGStorePF.reset_index(drop=True)

BGStorePF['Sales_$_LTD'] = BGStorePF['Sales_$_LTD']/1000
BGStorePF['Rcpt_$_LTD'] = BGStorePF['Rcpt_$_LTD']/1000

BGStorePF.replace(to_replace=0, value=np.nan, inplace=True)
BGStorePF.replace(to_replace='0', value=np.nan, inplace=True)

BGStorePF = BGStorePF.dropna(subset=['Product_Name'])
BGStorePF = BGStorePF.reset_index(drop=True)

BGStorePF['Store'] = 'bg'
BGStorePF['Type'] = 'store'
BGStorePF['Season_Code'] = np.nan
BGStorePF['group'] = np.nan

i = 0
for code in BGStorePF.Group_Code:
    if code == '0001_GROUP A':
        BGStorePF.loc[i, 'Season_Code'] = '16F'
        BGStorePF.loc[i, 'group'] = 'stock'
    elif code == '0002_GROUP B':
        BGStorePF.loc[i, 'Season_Code'] = '16F'
        BGStorePF.loc[i, 'group'] = 'consignment'
    elif code == '0003_GROUP C':
        BGStorePF.loc[i, 'Season_Code'] = 'PFP'
        BGStorePF.loc[i, 'group'] = 'stock'
    elif code == '0004_GROUP D':
        BGStorePF.loc[i, 'Season_Code'] = '16P'
        BGStorePF.loc[i, 'group'] = 'stock'
    elif code == '0005_GROUP E':
        BGStorePF.loc[i, 'Season_Code'] = '16P'
        BGStorePF.loc[i, 'group'] = 'consignment'
    elif code in [None, 0, np.nan, 'GC']:
        pass
    else:
        print 'Error value at row', i, ': ', code
        raise ValueError

    i = i + 1



BGOmniPF = pd.io.excel.read_excel(directory+BGDataPF, 2)
assert (BGOmniPF.loc[64][0] == 'PRE-FALL 17 EARLY'), 'Something\'s off with Prefall Stock, boss!'
assert (BGOmniPF.loc[146][0] == 'PRE-FALL 17 STOCK'), 'Something\'s off with Prefall Consignment, boss!'
assert (BGOmniPF.loc[310][0] == 'FALL 17 STOCK'), 'Something\'s off with Fall Stock, boss!'

BGOmniPF = BGOmniPF.drop(BGOmniPF.columns[[0,2,3,4,6,7,8,9,10,11,12,13,14,15,16,17,22,23,24,25,26,27,28,29]],axis=1)

BGOmniPF.columns = ['Group_Code','Product_Name','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD']
BGOmniPF.Product_Name.replace(to_replace='VENDOR STYLE', value=np.nan, inplace=True)

BGOmniPF = BGOmniPF.dropna(subset=['Product_Name'])
BGOmniPF = BGOmniPF.reset_index(drop=True)

BGOmniPF['Sales_$_LTD'] = BGOmniPF['Sales_$_LTD']/1000
BGOmniPF['Rcpt_$_LTD'] = BGOmniPF['Rcpt_$_LTD']/1000

BGOmniPF.replace(to_replace=0, value=np.nan, inplace=True)
BGOmniPF.replace(to_replace='0', value=np.nan, inplace=True)

BGOmniPF = BGOmniPF.dropna(subset=['Product_Name'])
BGOmniPF = BGOmniPF.reset_index(drop=True)

BGOmniPF['Store'] = 'bg'
BGOmniPF['Type'] = 'online'
BGOmniPF['Season_Code'] = np.nan
BGOmniPF['group'] = np.nan

i = 0
for code in BGOmniPF.Group_Code:
    if code == '0001_GROUP A':
        BGOmniPF.loc[i, 'Season_Code'] = '16F'
        BGOmniPF.loc[i, 'group'] = 'stock'
    elif code == '0002_GROUP B':
        BGOmniPF.loc[i, 'Season_Code'] = '16F'
        BGOmniPF.loc[i, 'group'] = 'consignment'
    elif code == '0003_GROUP C':
        BGOmniPF.loc[i, 'Season_Code'] = 'PFP'
        BGOmniPF.loc[i, 'group'] = 'stock'
    elif code == '0004_GROUP D':
        BGOmniPF.loc[i, 'Season_Code'] = '16P'
        BGOmniPF.loc[i, 'group'] = 'stock'
    elif code == '0005_GROUP E':
        BGOmniPF.loc[i, 'Season_Code'] = '16P'
        BGOmniPF.loc[i, 'group'] = 'consignment'
    elif code in [None, 0, np.nan, 'GC']:
        pass
    else:
        print 'Error value at row', i, ': ', code
        raise ValueError

    i = i + 1




# print 'NMOnline...'
# NMOnline = pd.io.excel.read_excel(directory+NMOnlineData, 0)
# assert (NMOnline.loc[4][1] == 'Style  Image'), 'Something\'s off with the alignment, boss!'

# NMOnline = NMOnline.drop(NMOnline.columns[[0,1,3,4,5,7,8,9,10,11,12,13,14,15,16,19,20,21,22,23,24,25,26,27,28,29,
#                                            30,31,32,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,
#                                            54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70]],axis=1)

# NMOnline.columns = ['Product_Name','Season_Code','Sales_$_LTD','Sales_U_LTD','OH_$']

# NMOnline['Store'] = 'nm'
# NMOnline['Type'] = 'online'

# NMOnline['Rcpt_$_LTD'] = NMOnline['OH_$'] + NMOnline['Sales_$_LTD']

# NMOnline.Season_Code.replace(to_replace='FY16 RESORT', value='16R', inplace=True)
# NMOnline.Season_Code.replace(to_replace='FY16 SPRING RUNWAY', value='16S', inplace=True)
# NMOnline.Season_Code.replace(to_replace='FY17 TRANSTN/PRE FALL', value='16P', inplace=True)
# NMOnline.Season_Code.replace(to_replace='FY17 FALL RUNWAY', value='16P', inplace=True)

# NMOnline = NMOnline.dropna(subset=['Product_Name'])



NMPrefall = pd.io.excel.read_excel(directory+NMPrefallData, 1, header = None)
test = NMPrefall[NMPrefall[0] == 'DETAIL STOP'].index.tolist()[0]
assert (NMPrefall.loc[0][17] == 'Regular'), 'These ain\'t TOTAL figures, boss!'
assert (NMPrefall.loc[test][0] == 'DETAIL STOP'), 'Something\'s off with the rows, boss!'

NMPrefall = NMPrefall.ix[8:test]
NMPrefall = NMPrefall.drop(NMPrefall.columns[[0,1,3,6,7,10,11,12,13,14,15,16,17,18]],axis=1)

NMPrefall.columns = ['test','OH_$','OH_$_LY','Sales_$_LTD','Sales_$_LTD_LY']

NMPrefall['Rcpt_$_LTD'] = NMPrefall['OH_$'].astype(float) + NMPrefall['Sales_$_LTD'].astype(float)
NMPrefall['Rcpt_$_LTD_LY'] = NMPrefall['OH_$_LY'].astype(float) + NMPrefall['Sales_$_LTD_LY'].astype(float)

NMPrefall['Season_Code'] = '16P'
NMPrefall['Store'] = 'nm'
NMPrefall['Type'] = 'store'

NMPrefall = NMPrefall.dropna(subset=['test'])
NMPrefall = NMPrefall[['Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','Season_Code','Store','Type']]



NMPFP = pd.io.excel.read_excel(directory+NMPFPData, 1, header = None)
test = NMPFP[NMPFP[0] == 'DETAIL STOP'].index.tolist()[0]
assert (NMPFP.loc[0][17] == 'Regular'), 'These ain\'t TOTAL figures, boss!'
assert (NMPFP.loc[test][0] == 'DETAIL STOP'), 'Something\'s off with the rows, boss!'

NMPFP = NMPFP.ix[8:test]
NMPFP = NMPFP.drop(NMPFP.columns[[0,1,3,6,7,10,11,12,13,14,15,16,17,18]],axis=1)

NMPFP.columns = ['test','OH_$','OH_$_LY','Sales_$_LTD','Sales_$_LTD_LY']

NMPFP['Rcpt_$_LTD'] = NMPFP['OH_$'].astype(float) + NMPFP['Sales_$_LTD'].astype(float)
NMPFP['Rcpt_$_LTD_LY'] = NMPFP['OH_$_LY'].astype(float) + NMPFP['Sales_$_LTD_LY'].astype(float)

NMPFP['Season_Code'] = 'PFP'
NMPFP['Store'] = 'nm'
NMPFP['Type'] = 'store'

NMPFP = NMPFP.dropna(subset=['test'])
NMPFP = NMPFP[['Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','Season_Code','Store','Type']]



NMSpecial = pd.io.excel.read_excel(directory+NMSpecialData, 1, header = None)
test = NMSpecial[NMSpecial[0] == 'DETAIL STOP'].index.tolist()[0]
assert (NMSpecial.loc[0][17] == 'Regular'), 'These ain\'t TOTAL figures, boss!'
assert (NMSpecial.loc[test][0] == 'DETAIL STOP'), 'Something\'s off with the rows, boss!'

NMSpecial = NMSpecial.ix[8:test]
NMSpecial = NMSpecial.drop(NMSpecial.columns[[0,1,3,6,7,10,11,12,13,14,15,16,17,18]],axis=1)
NMSpecial.columns = ['test','OH_$','OH_$_LY','Sales_$_LTD','Sales_$_LTD_LY']

NMSpecial['Rcpt_$_LTD'] = NMSpecial['OH_$'].astype(float) + NMSpecial['Sales_$_LTD'].astype(float)
NMSpecial['Rcpt_$_LTD_LY'] = NMSpecial['OH_$_LY'].astype(float) + NMSpecial['Sales_$_LTD_LY'].astype(float)

NMSpecial['Season_Code'] = 'special'
NMSpecial['Store'] = 'nm'
NMSpecial['Type'] = 'store'

NMSpecial = NMSpecial.dropna(subset=['test'])
NMSpecial = NMSpecial[['Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','Season_Code','Store','Type']]


print 'Saks...'
SFASeasons = pd.io.excel.read_excel(directory+SaksStoreData, 0, header = None)
assert (SFASeasons.loc[63][3] == 'Total'), 'Can\'t find the total, boss!'
prefall = SFASeasons[SFASeasons[3] == 'PREFALL'].index.tolist()[0]
fall = SFASeasons[SFASeasons[3] == 'FALL'].index.tolist()[0]

SFASeasons = SFASeasons.ix[9:120]
SFASeasons = SFASeasons.drop(SFASeasons.columns[[0,1,3,4,6,8,11,12,13,14,15,16,17,18,19,20,21,22]],axis=1)
SFASeasons.columns = ['City_Code','Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','Season_Code']

SFASeasons['Season_Code'][prefall] = '16P'
SFASeasons['Season_Code'][fall] = '16F'
SFASeasons['Season_Code'].fillna(method='ffill', inplace=True)
SFASeasons['Store'] = 'sfa'
SFASeasons['Type'] = 'store'

SFASeasons = SFASeasons[SFASeasons.City_Code != 'CHECK']
SFASeasons = SFASeasons[SFASeasons.Season_Code != 'X']
SFASeasons = SFASeasons.dropna(subset=['City_Code'])
SFASeasons = SFASeasons[['City_Code','Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY','Season_Code','Store','Type']]
SFASeasons['OH_$'] = np.nan
SFASeasons['Sales_$_WTD'] = np.nan
SFASeasons['Report_Date'] = np.nan




# SFASpecial = pd.io.excel.read_excel(directory+SaksStoreData, 2, header = None)
# print SFASpecial.head(15)

# date = SFASpecial.loc[1][22]
# # date = dparser.parse(text, fuzzy=True).date()

# SFASpecial = SFASpecial[[2,5,7,9,10]]
# SFASpecial.columns = ['test','Sales_$_LTD','Sales_$_LTD_LY','Rcpt_$_LTD','Rcpt_$_LTD_LY']

# SFASpecial = SFASpecial.dropna(subset=['City_Code'])
# SFASpecial = SFASpecial.reset_index(drop=True)

# SFASpecial['Store'] = 'sfa'
# SFASpecial['Report_Date'] = date
# SFASpecial['Season_Code'] = 'special'

# SFASpecial = SFASpecial.ix[250:]
# dataRow = np.where(SFASpecial['test'] == 'CHECK')[0][0]
# print SFASpecial.head()
# SFASpecial = SFASpecial.ix[dataRow:]
# SFASpecial = SFASpecial.reset_index(drop=True)
# print SFASpecial.head()

# SFASpecial['Sales_$_LTD'] = SFASpecial['Sales_$_LTD'] / 1000
# SFASpecial['Rcpt_$_LTD'] = SFASpecial['Rcpt_$_LTD'] / 1000

# SFASpecial = SFASpecial.drop('test', axis=1)


StoreSalesDB = pd.merge(BGStorePF, pd.merge(BGOmniPF, pd.merge(NMPrefall, pd.merge(NMSpecial, pd.merge(SFASeasons, NMPFP, how='outer'), how='outer'), how='outer'), how='outer'), how='outer')
StoreSalesDB = StoreSalesDB.reset_index(drop=True)

StoreSalesDB = StoreSalesDB[['Season_Code','group','Product_Name','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD','Store','Type','OH_$','Sales_$_LTD_LY','Rcpt_$_LTD_LY','City_Code','Sales_$_WTD','Report_Date']]

i = 0
for name in StoreSalesDB.Product_Name:
    if (name == '3027GDL') | (name == '3027GQI') | (name == '3037GDL') | (name == '3037GQI') | (name == '3973GQJ') | (name == '4003GQI') |\
    (name == '4003GDL') | (name == '4063GDL') | (name == '5027GDL') | (name == '5037GDL') | (name == '5077GDL') | (name == '5483GQI') |\
    (name == '5633GDL') | (name == '5633GQI') | (name == '5793GQJ') | (name == '5803GQJ'):
        StoreSalesDB.loc[i, 'Season_Code'] = 'PFP'
    i = i + 1

print 'Merged Store sales data! \n'

"""
END Store Sales


















BEGIN Style Sales
"""
    

BGStylesPF = pd.io.excel.read_excel(directory+BGDataPF, header=None)

lastRow = np.where(BGStylesPF[0] == 'BASICS')[0][0]
assert (BGStylesPF.loc[1][0] == 'WEEKLY DETAILED SELLING'), 'Paste job\'s off, boss!'
assert (BGStylesPF.loc[5][2] == 'SKU'), 'Columns look off, boss! --> SKU'
assert (BGStylesPF.loc[5][26] == 'OO U'), 'Columns look off, boss! --> OO U'

BGStylesPF = BGStylesPF.drop(BGStylesPF.columns[[0,2,3,4,6,7,10,11,14,15,16,17,22,23,28,29]],axis=1)
BGStylesPF = BGStylesPF.ix[6:lastRow+1]
BGStylesPF.columns = ['Group_Code','Product_Name','Color_Desc','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD','OH_U','OH_$','OO_U','OO_$']

BGStylesPF = BGStylesPF.replace(to_replace='0', value=np.nan)
BGStylesPF = BGStylesPF.replace(to_replace=0, value=np.nan)
BGStylesPF['Store'] = 'bg'
BGStylesPF['Season_Code'] = np.nan

BGStylesPF = BGStylesPF.dropna(subset=['Product_Name'])
BGStylesPF = BGStylesPF.reset_index(drop=True)

i = 0
for code in BGStylesPF.Group_Code:
    if code == 'GROUP A':
        BGStylesPF.loc[i, 'Season_Code'] = '16F'
    elif code == 'GROUP B':
        BGStylesPF.loc[i, 'Season_Code'] = '16F'
    elif code == 'GROUP C':
        BGStylesPF.loc[i, 'Season_Code'] = 'PFP'
    elif code == 'GROUP D':
        BGStylesPF.loc[i, 'Season_Code'] = '16P'
    elif code == 'GROUP E':
        BGStylesPF.loc[i, 'Season_Code'] = '16P'
    elif code in [None, 0, np.nan]:
        pass
    else:
        print 'Error value at row', i, ': ', code
        raise ValueError

    i = i + 1

BGStylesPF['Sales_$_LTD'] = BGStylesPF['Sales_$_LTD'].astype(float)/1000
BGStylesPF['Rcpt_$_LTD'] = BGStylesPF['Rcpt_$_LTD'].astype(float)/1000
BGStylesPF['Sales_$_WTD'] = BGStylesPF['Sales_$_WTD'].astype(float)/1000
BGStylesPF['OH_$'] = BGStylesPF['OH_$'].astype(float)/1000
BGStylesPF['OO_$'] = BGStylesPF['OO_$'].astype(float)/1000

i = 0
for color in BGStylesPF.Color_Desc:
    BGStylesPF.loc[i, 'Color_Desc'] = str(color[4:])
    i += 1



# print 'NMStyles...'
# NMOnlineStyles = pd.io.excel.read_excel(directory+NMOnlineData, 0)
# assert (NMOnlineStyles.loc[4][1] == 'Style  Image'), 'Something\'s off with the alignment, boss!'
# assert (NMOnlineStyles.loc[4][70] == 'Current Inventory Status'), 'Column number\'s off, boss!'

# NMOnlineStyles = NMOnlineStyles.ix[5:]
# NMOnlineStyles = NMOnlineStyles.drop(NMOnlineStyles.columns[[0,1,3,4,5,7,8,12,13,14,15,16,19,20,21,22,23,24,25,26,27,28,29,
#                                            30,31,34,35,36,37,38,39,40,41,42,43,44,45,48,49,50,51,52,53,
#                                            54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70]],axis=1)

# NMOnlineStyles.columns = ['Product_Name','Season_Code','Orig_Retail_Price','Sales_$_WTD','Sales_U_WTD','Sales_$_LTD','Sales_U_LTD','OH_$','OH_U',
#                    'OO_$','OO_U']

# NMOnlineStyles['Store'] = 'nm'
# NMOnlineStyles['Rcpt_$_LTD'] = NMOnlineStyles['OH_$'] + NMOnlineStyles['Sales_$_LTD']
# NMOnlineStyles['Rcpt_U_LTD'] = NMOnlineStyles['OH_U'] + NMOnlineStyles['Sales_U_LTD']
# NMOnlineStyles.Season_Code.replace(to_replace='FY16 RESORT', value='16R', inplace=True)
# NMOnlineStyles.Season_Code.replace(to_replace='FY16 SPRING RUNWAY', value='16S', inplace=True)
# NMOnlineStyles.Season_Code.replace(to_replace='FY17 TRANSTN/PRE FALL', value='16P', inplace=True)
# NMOnlineStyles.Season_Code.replace(to_replace='FY17 FALL RUNWAY', value='16P', inplace=True)

# NMOnlineStyles.replace(to_replace='0', value=np.nan, inplace=True)
# NMOnlineStyles.replace(to_replace=0, value=np.nan, inplace=True)

# NMOnlineStyles = NMOnlineStyles.dropna(subset=['Product_Name'])
# NMOnlineStyles = NMOnlineStyles.reset_index(drop=True)



NMStoreStyles = pd.io.excel.read_excel(directory+NMStyleData, 0, header = None)
assert (NMStoreStyles.loc[0][8] == 'Regular'), 'These ain\'t TOTAL figures, boss!'
assert (NMStoreStyles.loc[8][29] == 'SHARED'), 'Something\'s off with the alignment, boss!'

NMStoreStyles = NMStoreStyles.ix[9:]
NMStoreStyles = NMStoreStyles.drop(NMStoreStyles.columns[[0,1,2,3,4,5,7,8,9,10,12,15,22,25,28,29,30]],axis=1)

NMStoreStyles.columns = ['Product_Name','rawseason','Color_Desc','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','Rcpt_U_LTD','Rcpt_$_LTD','OO_U','OO_$']
NMStoreStyles['Store'] = 'nm'
NMStoreStyles['Sales_$_LTD'] = NMStoreStyles['Sales_$_LTD'].astype(float)/1000
NMStoreStyles['Sales_$_WTD'] = NMStoreStyles['Sales_$_WTD'].astype(float)/1000
NMStoreStyles['Rcpt_$_LTD'] = NMStoreStyles['Rcpt_$_LTD'].astype(float)/1000
NMStoreStyles['OH_$'] = NMStoreStyles['OH_$'].astype(float)/1000
NMStoreStyles['OO_$'] = NMStoreStyles['OO_$'].astype(float)/1000

NMStoreStyles = NMStoreStyles.replace(to_replace='0', value=np.nan)
NMStoreStyles = NMStoreStyles.replace(to_replace=0, value=np.nan)

NMStoreStyles = NMStoreStyles.dropna(subset=['Product_Name'])
NMStoreStyles = NMStoreStyles.reset_index(drop=True)

NMStoreStyles['Season_Code'] = ''
i = 0
for item in NMStoreStyles.rawseason:
    if 'S16' in item:
        NMStoreStyles.loc[i, 'Season_Code'] = '16S'
    elif 'R16' in item:
        NMStoreStyles.loc[i, 'Season_Code'] = '16R'
    elif 'PF17' in item:
        NMStoreStyles.loc[i, 'Season_Code'] = '16P'
    elif 'F17' in item:
        NMStoreStyles.loc[i, 'Season_Code'] = '16P'
    elif 'EPF' in item:
        NMStoreStyles.loc[i, 'Season_Code'] = '16P'
    i = i + 1

NMStoreStyles = NMStoreStyles[['Product_Name','Color_Desc','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','Rcpt_U_LTD','Rcpt_$_LTD','OO_U','OO_$','Store']]



print 'Saks styles...'
SaksStyles = pd.io.excel.read_excel(directory+SaksStyleData, 0, header=None)

firstRow = np.where(SaksStyles[0] == 'Delivery')[0][0]
if SaksStyles.loc[firstRow][2] == 'Class':
    assert (SaksStyles.loc[firstRow][23] == 'OO U'), 'Something\'s off with the alignment, boss!'
    SaksStyles = SaksStyles.drop(SaksStyles.columns[[0,2,3,4,6,7,8,10,15,16,17,18,25,26,29,30,31,32]],axis=1)
elif SaksStyles.loc[firstRow][2] == 'Vendor Style':
    assert (SaksStyles.loc[firstRow][21] == 'OO U'), 'Something\'s off with the alignment, boss!'
    SaksStyles = SaksStyles.drop(SaksStyles.columns[[0,2,4,5,6,8,13,14,15,16,23,24,27,28,29,30]],axis=1)
else:
    print SaksStyles.head(15)
    raise ValueError

SaksStyles.columns = ['Season_Temp','Product_Name','Color_Desc','Last_Received','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','OH_U','OH_$','OO_U','OO_$','Rcpt_U_LTD','Rcpt_$_LTD']

SaksStyles['Store'] = 'sfa'
SaksStyles['Product_Name'] = SaksStyles['Product_Name'].fillna(method='ffill')
SaksStyles['Season_Temp'] = SaksStyles['Season_Temp'].fillna(method='ffill')

SaksStyles = SaksStyles.ix[12:]
SaksStyles = SaksStyles.reset_index(drop=True)

SaksStyles['Sales_$_LTD'] = SaksStyles['Sales_$_LTD'].astype(float)/1000
SaksStyles['Rcpt_$_LTD'] = SaksStyles['Rcpt_$_LTD'].astype(float)/1000
SaksStyles['Sales_$_WTD'] = SaksStyles['Sales_$_WTD'].astype(float)/1000
SaksStyles['OH_$'] = SaksStyles['OH_$'].astype(float)/1000
SaksStyles['OO_$'] = SaksStyles['OO_$'].astype(float)/1000

SaksStyles = SaksStyles.replace(to_replace='0', value=np.nan)
SaksStyles = SaksStyles.replace(to_replace=0, value=np.nan)
SaksStyles.Color_Desc = SaksStyles.Color_Desc.str.upper()
SaksStyles.Color_Desc = SaksStyles.Color_Desc.str.replace('OPEN ', '')

SaksStyles = SaksStyles.dropna(subset=['Product_Name'])
SaksStyles = SaksStyles.dropna(subset=['Color_Desc'])
SaksStyles = SaksStyles.reset_index(drop=True)

SaksStyles['Season_Code'] = np.nan
i = 0
for code in SaksStyles.Season_Temp:
    if code == 'Pre-Fall 2016':
        SaksStyles.loc[i, 'Season_Code'] = '16P'
    elif code == 'Fall 2016':
        SaksStyles.loc[i, 'Season_Code'] = '16F'
    elif code == '51CON':
        SaksStyles.loc[i, 'Season_Code'] = 'ICON'
    elif code in [None, 0, np.nan]:
        pass
    else:
        print 'Error value at row', i, ': ', code
        raise ValueError
    i = i + 1


StyleSalesDB = pd.merge(BGStylesPF, pd.merge(NMStoreStyles, SaksStyles, how='outer'), how='outer')
StyleSalesDB = StyleSalesDB.dropna(subset=['Store'])
StyleSalesDB = StyleSalesDB.reset_index(drop=True)

i = 0
for name in StyleSalesDB.Product_Name:
    if (name == '3027GDL') | (name == '3027GQI') | (name == '3037GDL') | (name == '3037GQI') | (name == '3973GQJ') | (name == '4003GQI') |\
    (name == '4003GDL') | (name == '4063GDL') | (name == '5027GDL') | (name == '5037GDL') | (name == '5077GDL') | (name == '5483GQI') |\
    (name == '5633GDL') | (name == '5633GQI') | (name == '5793GQJ') | (name == '5803GQJ'):
        StyleSalesDB.loc[i, 'Season_Code'] = 'PFP'
    i = i + 1

i = 0
StyleSalesDB['classes'] = ''
for item in StyleSalesDB['Product_Name']: 
    if type(item) != unicode:
        StyleSalesDB.loc[i, 'classes'] = str(item)[:1]
    else:
        temp = item[:2]
        if temp == "AP":
            StyleSalesDB.loc[i, 'classes'] = item[3:4]
        else:
            StyleSalesDB.loc[i, 'classes'] = item[:1]
    i = i + 1


"""
END Style Sales

















BEGIN Specialty functions
"""
print 'Specialty information...'
colors = StyleSalesDB[['Season_Code','Color_Desc','Sales_U_LTD']]
colors = colors[(colors.Season_Code == '16S') | (colors.Season_Code == '16F') | (colors.Season_Code == '16P') | (colors.Season_Code == 'PFP')]
colors = colors.groupby([colors['Color_Desc']])['Sales_U_LTD'].sum().sort_values(ascending=False)[:15]

fabrics = []
for fabric in StyleSalesDB['Product_Name']:
    fabrics.append(fabric[4:])
StyleSalesDB['fabric'] = fabrics
fabrics = StyleSalesDB[['Season_Code','fabric','Sales_U_LTD']]
fabrics = fabrics[(fabrics.Season_Code == '16S') | (fabrics.Season_Code == '16F') | (fabrics.Season_Code == '16P') | (fabrics.Season_Code == 'PFP')]
fabrics = fabrics.groupby([fabrics.fabric])['Sales_U_LTD'].sum().sort_values(ascending=False)[:15]

topPrefall = StyleSalesDB[['Season_Code','Product_Name','Sales_U_LTD','Sales_$_LTD']][(StyleSalesDB.Season_Code == '16P')]
topPrefall.Sales_U_LTD = pd.to_numeric(topPrefall.Sales_U_LTD, errors='coerce')
topPrefall['Sales_$_LTD'] = pd.to_numeric(topPrefall['Sales_$_LTD'], errors='coerce')
topPrefall = topPrefall.groupby([topPrefall['Product_Name']]).sum()
topPrefall = topPrefall.sort_values(by=['Sales_U_LTD','Sales_$_LTD'], ascending=False)[:5]
topPrefall = topPrefall.reset_index()
i = 0
topPrefall['Class_Indicator'] = np.nan
for style in topPrefall.Product_Name.values:
    topPrefall.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1

topPFP = StyleSalesDB[['Season_Code','Product_Name','Sales_U_LTD','Sales_$_LTD']][(StyleSalesDB.Season_Code == 'PFP')]
topPFP.Sales_U_LTD = pd.to_numeric(topPFP.Sales_U_LTD, errors='coerce')
topPFP['Sales_$_LTD'] = pd.to_numeric(topPFP['Sales_$_LTD'], errors='coerce')
topPFP = topPFP.groupby([topPFP['Product_Name']]).sum()
topPFP = topPFP.sort_values(by=['Sales_U_LTD','Sales_$_LTD'], ascending=False)[:5]
topPFP = topPFP.reset_index()
i = 0
topPFP['Class_Indicator'] = np.nan
for style in topPFP.Product_Name.values:
    topPFP.loc[i, 'Class_Indicator'] = style[0:1]
    i = i + 1
    

StyleSalesDB = StyleSalesDB[['Season_Code','Product_Name','Color_Desc','Orig_Retail_Price','Sales_U_WTD','Sales_$_WTD','Sales_U_LTD','Sales_$_LTD','Rcpt_U_LTD','Rcpt_$_LTD','OH_U','OH_$','OO_U','OO_$','Store','Last_Received','classes','fabric']]


StoreSalesDB.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\StoreSalesDB.csv')
StyleSalesDB.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\StyleSalesDB.csv')
colors.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\colors.csv')
fabrics.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\\fabrics.csv')
topPrefall.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\\topPrefall.csv', index=False)
topPFP.to_csv('U:\Retail Reporting\AppFolder\Wholesale\Out\YummyData\\topPFP.csv', index=False)

dealWithFiles()

print 'All done!'