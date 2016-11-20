import pandas as pd
import numpy as np

def combColors(df):

	attrMap = 'U:\Retail Reporting\MagicFolder\Retail\In\AttrMap.xlsx'

	colorDict = pd.io.excel.read_excel(attrMap, 0)
	colorfill = pd.io.excel.read_excel(attrMap, 5)
	singlecolor = pd.io.excel.read_excel(attrMap, 6)

	colorDict = dict(zip(colorDict['Key'].str.lower(), colorDict['Value'].str.lower()))
	colorfill = dict(zip(colorfill['Key'].astype(str), colorfill['Value'].str.lower()))
	singlecolor = dict(zip(singlecolor['Key'].astype(str), singlecolor['Value'].str.lower()))


	j = 0
	for i in df.Color_Desc.values:
		if pd.isnull(i):
			tempStyle = df.Product_Name[j]
			tempColor = singlecolor.get(tempStyle, i)
			# if pd.isnull(tempColor):
			# 	tempColor = colorfill.get(tempStyle, i)
			df.loc[j, 'Color_Desc'] = tempColor
		j = j + 1
	
	df['style-color'] = np.nan
	df['style-color'] = df.Product_Name.str.cat(df.Color_Desc, sep="-")

	df.Color_Desc = df['style-color'].str.lower().map(colorfill)
	df = df.reset_index(drop=True)

	return df
Product_Name','Color_Desc