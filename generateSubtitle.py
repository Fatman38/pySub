#don't forget to install the following
# pip install pandas
# pip install openpyxl

# usage:
# copy excel file in the same folder as this script AND rename to: "inputs.xlsx"
#		rename all its column by its country code and rename the Time column by "Key"
# copy each base str files in the same folder as this script AND rename each of them to "MODULE X.srt"
# launch generateSubtitle.py

import pandas as pd
from pathlib import Path
from itertools import islice
import numpy as np

def getDfFrom(filename):
	myfile = open(filename, "r")
	indexes = []
	timestamps = []

	lines = myfile.readlines()
	c = 0;
	for line in lines:
		if c == 0:
			indexes.append(line.strip())
		elif c == 1:
			timestamps.append(line.strip())
		c = c + 1
		if c > 3:
			c = 0
	return pd.DataFrame(timestamps,indexes, columns = ["ts"])

def createSubtitleFile(timestamps, translation, column, modulename):
	outputfilename = column + "-" + modulename + ".srt"
	print(outputfilename)
	
	df2 = translation

	timestamps['Key']=timestamps['Key'].astype(str)
	df2['Key']=df2['Key'].astype(str)

	df3 = timestamps.merge(df2, how="outer", on='Key')
	print(df3)
	
	outputfile = open(outputfilename, "w", encoding='utf-8')
	for index, row in df3.iterrows():
		text = ""
		if not pd.isnull(row['Key']):
			text = str(row['Key'])
		outputfile.write(text + "\n")
		text = ""
		if not pd.isnull(row['ts']):
			text = str(row['ts'])
		outputfile.write(text + "\n")
		text = ""
		if not pd.isnull(row[column]):
			text = str(row[column])
		outputfile.write(text + "\n")
		outputfile.write("\n")
		
# ++++++++++++++++++++++++++++++
# ++++++++++++ MAIN ++++++++++++ 
# ++++++++++++++++++++++++++++++

xl = pd.ExcelFile('inputs.xlsx')

print (xl.sheet_names)

for sheet in xl.sheet_names:
	sub_filename = sheet + ".srt"
	print("sub: ", sub_filename)

	sub_file = Path(sub_filename)
	if not sub_file.is_file():
		continue

	sdf = getDfFrom(sub_filename)
	sdf['Key'] = sdf.index
	print(sdf)

	df = pd.read_excel(xl, sheet_name=sheet)
	df.dropna(subset = ['Key'], inplace=True)
	print (df)
	if (df['Key'].dtypes != int):
		df['Key'] = df['Key'].astype(np.int64)

	newindex = df[df.columns[0]]
	print(newindex)
	df2 = df
	df2 = df2.set_index(newindex)
	print(df2)
	print()
	print()
	
	for column in df2:
		if column != 'Key':
			createSubtitleFile(sdf, df[['Key', column]], column, sheet)
		
