import pandas as pd
import numpy as np

def standardizeDTypes(df):

	df['Store_Name'] = df['Store_Name'].apply(str)
	df['Vendor_Code'] = df['Vendor_Code'].apply(str)
	df['Season_Code'] = df['Season_Code'].apply(str)
	df['Sub_Dept'] = df['Sub_Dept'].apply(str)
	df['Product_Name'] = df['Product_Name'].apply(str)
	df['Color_Desc'] = df['Color_Desc'].apply(str)
	df['Size_Code'] = df['Size_Code'].apply(str)
	df['Fabric'] = df['Fabric'].apply(str)

	df['Orig_Retail_Price'] = df['Orig_Retail_Price'].apply(float)
	df['mtdunits_16'] = df['mtdunits_16'].apply(float)
	df['mtd$_16'] = df['mtd$_16'].apply(float)
	df['mtdunits_15'] = df['mtdunits_15'].apply(float)
	df['mtd$_15'] = df['mtd$_15'].apply(float)
	df['ytdunits_16'] = df['ytdunits_16'].apply(float)
	df['ytd$_16'] = df['ytd$_16'].apply(float)
	df['ytdunits_15'] = df['ytdunits_15'].apply(float)
	df['ytd$_15'] = df['ytd$_15'].apply(float)
	df['ltdunits_16'] = df['ltdunits_16'].apply(float)
	df['ltd$_16'] = df['ltd$_16'].apply(float)
	df['ltdunits_15'] = df['ltdunits_15'].apply(float)
	df['ltd$_15'] = df['ltd$_15'].apply(float)
	df['rcptunits_15'] = df['rcptunits_15'].apply(float)
	df['rcpt$_15'] = df['rcpt$_15'].apply(float)
	df['rcptunits_16'] = df['rcptunits_16'].apply(float)
	df['rcpt$_16'] = df['rcpt$_16'].apply(float)
	df['Onhand_Qty'] = df['Onhand_Qty'].apply(float)
	df['Onhand_Retail'] = df['Onhand_Retail'].apply(float)
	df['Class_Num'] = df['Class_Num'].apply(float)
	df['Month'] = df['Month'].apply(float)

	return df

def stripWhitespace(df):

	df['Store_Name'] = df['Store_Name'].str.strip()
	df['Vendor_Code'] = df['Vendor_Code'].str.strip()
	df['Season_Code'] = df['Season_Code'].str.strip()
	df['Sub_Dept'] = df['Sub_Dept'].str.strip()
	df['Product_Name'] = df['Product_Name'].str.strip()
	df['Color_Desc'] = df['Color_Desc'].str.strip()
	df['Size_Code'] = df['Size_Code'].str.strip()
	df['Fabric'] = df['Fabric'].str.strip()

	return df

def replaceUnwantedValues(df):

	df.Store_Name.replace(to_replace="CHNY Madison Ave.", value="MAD", inplace=True)
	df.Store_Name.replace(to_replace="CHNY Rodeo Drive", value="LA", inplace=True)
	df.Store_Name.replace(to_replace="CHNY Melrose Pl.", value="LA", inplace=True)
	df.Store_Name.replace(to_replace="CHNY Dallas", value="DAL", inplace=True)
	df.Store_Name.replace(to_replace="CHNY Harrods", value="HAR", inplace=True)

	df.Size_Code.replace(to_replace="LG", value="L", inplace=True)

	df.Product_Name = df.Product_Name.str.replace(pat=" ", repl="")

	df.Vendor_Code.replace(to_replace='CHNY', value='CH', inplace=True)

	df.Season_Code.replace(to_replace='16S', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='16R', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='SP16', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='CR16', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='RE16', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='FA16', value='16F', inplace=True)
	df.Season_Code.replace(to_replace='PF16', value='16P', inplace=True)
	df.Season_Code.replace(to_replace='SP15', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='CR15', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='15S', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='15R', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='FA15', value='15F', inplace=True)
	df.Season_Code.replace(to_replace='PF15', value='15P', inplace=True)

	df.Season_Code.replace(to_replace='MTO', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='00C', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='FA13', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='FA12', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='CR14', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='CR13', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='PF12', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='PF13', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='SP14', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='SP13', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='SP11', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='FA14', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='PF14', value='COVR', inplace=True)

	df.Sub_Dept.replace(to_replace='SHOE', value='SHOES', inplace=True)


	i = 0
	for name in df.Product_Name:
	    if (name in ['3027GDL','3027GQI','3037GDL','3037GQI','3973GQJ','4003GQI','4003GDL','4063GDL','5027GDL','5037GDL','5077GDL','5483GQI','5633GDL','5633GQI','5793GQJ','5803GQJ']) & (df.loc[i, 'Season_Code'] == '16P'):
	        
	        df.loc[i, 'Season_Code'] = 'PFP'

	    elif (name in ['201P','208P','209P','204P','2057ZNC','2007CDC','2017CDC','2160CJM','2027GDK','2017GDK','2037GDK','2047LAC','2040ATAEMB','2010PTS','2140TEM','003P','001P','002P','0010SWTEMB','1007LEA','1027LEA','1017LAC','1010PTS','1040PTS','3037LBA','3009RAO','3027DNM','3027COT','3037LEA','3027LEA','3057JTX','3007LAC','3020PTS','3040PTS','3030PTS','3050CJT','3080MAU','4003LBA','4013LBA','4023LBA','4007COT','4030CPY','4040CPY','5087DNM','5097DNM','5009RAY','5039RAO','5029RAO','5037COT','5047GDK','5057COT','5057SLK','5080MAU','5150MAU','5017LAC','5007LAC','5027LAC','5160FIN','5140PTSEMB','7040CJT','7050CCC','7010CJA','7110MAU','7070RCC','7020PTS','7030PTSEMB','7120PTS','9017BLT','9017NG','9017GDH','9007BLT']) & (df.loc[i, 'Season_Code'] in ['CPC','DAC','ESC','EVC']):

	    	df.loc[i, 'Season_Code'] = 'ICON'

	    elif name in ['SORTW','SOBRIDAL','SOFUR','SORTWCR14','SORTWSP14','SORTWSP15','SORTWCPC','SOFUR123796']:
	        df.loc[i, 'Season_Code'] = 'SPO'
	        df.loc[i, 'Orig_Retail_Price'] = 'N/A'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'

	    elif name in ['RTWUPCHARGE','UPCHARGEMTO','UPCHARGEFA14','UPCHARGEFA12']:
	        df.loc[i, 'Season_Code'] = 'UPCH'
	        df.loc[i, 'Sub_Dept'] = 'UPCHARGE'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'
	        df.loc[i, 'Orig_Retail_Price'] = 'N/A'

	    elif name in ['INVALIDRTW','INVALIDBRD']:
	        df.loc[i, 'Season_Code'] = 'INVALID'
	        df.loc[i, 'Sub_Dept'] = 'INVALID'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'
	        df.loc[i, 'Orig_Retail_Price'] = 'N/A'

	    elif name in ['BRDUPCHARGE','UPCHARGE1','UPCHARGEBRIDAL','UPCHARGEBRIDAL1','UPCHARGEBRIDAL2','UPCHARGEBRIDAL3']:
	        df.loc[i, 'Season_Code'] = 'UPCH'
	        df.loc[i, 'Vendor_Code'] = 'BRIDAL'
	        df.loc[i, 'Sub_Dept'] = 'UPCHARGE'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'
	        df.loc[i, 'Orig_Retail_Price'] = 'N/A'

	    elif name in ['BRIDALSCARF']:
	        df.loc[i, 'Season_Code'] = 'COVR'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Orig_Retail_Price'] = 'N/A'

	    elif name == 'VCHD':
	    	df.loc[i, 'Season_Code'] = 'SP05'

	    elif name in ['WISPY', 'WISPY2']:
	        df.loc[i, 'Size_Code'] = 'OS'


	    i = i + 1

	i = 0
	for dept in df.Sub_Dept:
	    if dept in ['JEWELRY','BRACELET','EARRINGS']:
	        df.loc[i, 'Season_Code'] = 'JEWELRY'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'

	    elif dept in ['FUR']:
	        df.loc[i, 'Season_Code'] = 'COVR'

	    elif dept in ['LARGECANDLE','CANDLEFORDISPLAY','MINICANDLE','CANDLE','PERFUME']:
	        df.loc[i, 'Season_Code'] = 'FRAGRANCE'
	        df.loc[i, 'Color_Desc'] = 'N/A'
	        df.loc[i, 'Size_Code'] = 'N/A'

	    elif dept in ['SHOES']:
	        df.loc[i, 'Season_Code'] = 'SHOES'

	    elif dept in ['HANDBAG']:
	        df.loc[i, 'Season_Code'] = 'HANDBAG'
	        df.loc[i, 'Size_Code'] = 'N/A'

	    elif dept in ['SUNGLASSES']:
	        df.loc[i, 'Season_Code'] = 'SUNGLASSES'
	        df.loc[i, 'Size_Code'] = 'N/A'

	    elif dept in ['PIN','BOOK']:
	        df.loc[i, 'Season_Code'] = 'COVR'
	        df.loc[i, 'Size_Code'] = 'N/A'
	        df.loc[i, 'Color_Desc'] = 'N/A'

	    i = i + 1

	i = 0
	for vendor in df.Vendor_Code:
	    if vendor == 'TOSCHI':
	        df.loc[i, 'Season_Code'] = 'COVR'
	        df.loc[i, 'Sub_Dept'] = 'FUR'

	    elif vendor in ['TYLER','TYLERC']:
	        df.loc[i, 'Season_Code'] = 'HANDBAG'
	        df.loc[i, 'Sub_Dept'] = 'HANDBAG'
	        df.loc[i, 'Size_Code'] = 'N/A'

	    i = i + 1

	df.Season_Code.replace(to_replace='CPC', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='DAC', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='ESC', value='COVR', inplace=True)
	df.Season_Code.replace(to_replace='EVC', value='COVR', inplace=True)

	return df

def filterForStores(df):

	df = df[(df.Store_Name.isin(['HAR','MAD','DAL','LA','CHNY Replenishment WHS']))]
	df = df.reset_index(drop=True)

	return df

def cleanColors(df):

	attrMap = 'U:\Retail Reporting\MagicFolder\Retail\In\AttrMap.xlsx'

	colorfill = pd.io.excel.read_excel(attrMap, 5)
	singlecolor = pd.io.excel.read_excel(attrMap, 6)

	colorfill = dict(zip(colorfill['Key'].astype(str), colorfill['Value'].str.lower()))
	singlecolor = dict(zip(singlecolor['Key'].astype(str), singlecolor['Value'].str.lower()))

	df['style-color'] = df['Product_Name'].str.cat(df['Color_Desc'], sep='-')
	df.Color_Desc = df['style-color'].map(colorfill)

	j = 0
	for col in df.Color_Desc.values:
		if pd.isnull(col):
			tempStyle = df.Product_Name[j]
			tempColor = singlecolor.get(tempStyle, col)
			if pd.isnull(tempColor):
				tempStyle = df['style-color'][j]
				tempColor = colorfill.get(tempStyle, col)
				if pd.isnull(tempColor):
					tempColor = 'unknown'
			df.loc[j, 'Color_Desc'] = tempColor
		j = j + 1

	df = df.drop('style-color', axis=1)

	return df

def fillRetailPrices(df):

	attrMap = 'U:\Retail Reporting\MagicFolder\Retail\In\AttrMap.xlsx'

	priceDict = pd.io.excel.read_excel(attrMap, 1)

	priceDict = dict(zip(priceDict['Key'].astype(str), priceDict['Value']))

	df.Orig_Retail_Price = df['Product_Name'].map(priceDict)

	return df

def cleanProductNames(df):

	styles = []
	for style in df['Product_Name'].values:
	    try:
	        cleaned = style.split('-')
	        if len(cleaned) == 2:
	            styles.append(cleaned[1])
	        else:
	            styles.append(cleaned[0])
	    except AttributeError:
	        styles.append(style)
	        continue
	df['Product_Name'] = styles

	return df

def extractFabrics(df):

	fabrics = []
	for style in df['Product_Name']:
	    try:
	        fabrics.append(style[4:])
	    except TypeError:
	        fabrics.append(style)
	        continue
	df['Fabric'] = fabrics

	df.Fabric = df.Fabric.astype(str)

	return df

def extractClassTypes(df):

	i = 0
	for item in df['Product_Name']:
	    temp = item[:2]
	    if temp == 'AP':
	        df.loc[i, 'Class_Num'] = item[3:4]
	    else:
	        df.loc[i, 'Class_Num'] = item[:1]
	    i = i + 1
	df.Class_Num = df.Class_Num.astype(str)

	return df

def processData(df):
	
	df = standardizeDTypes(df)

	df = stripWhitespace(df)
	
	df = cleanProductNames(df)

	df = extractFabrics(df)

	df = extractClassTypes(df)

	df = cleanColors(df)

	df = replaceUnwantedValues(df)

	df = filterForStores(df)
	
	df = fillRetailPrices(df)

	return df