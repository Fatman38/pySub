#don't forget to install the following
# pip install pandas
# pip install openpyxl

# usage:
# copy excel file in the same folder as this script AND rename to: "inputs.xlsx"
# copy each base str files in the same folder as this script AND rename each of them to "MODULE X.srt"
# launch generateSubtitle.py

import pandas as pd
from pathlib import Path
from itertools import islice

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
	#df = timestamps
	
	#print(df)
	timestamps['Key'] = timestamps.index
	df2 = translation.to_frame()
	df2['Key'] = df2.index
	print(df2)
	timestamps['Key']=timestamps['Key'].astype(int)
	df2['Key']=df2['Key'].astype(int)

	df3 = timestamps.merge(df2, how="outer", on='Key')
	print(df3)
	
	outputfile = open(outputfilename, "w")
	for index, row in df3.iterrows():
		outputfile.write(str(row['Key']) + "\n")
		outputfile.write(str(row['ts']) + "\n")
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
	print (df)
	df2 = df.groupby(df.columns[0]).agg(lambda x : ' '.join(x))
	print (df2)
	print()
	print()
	
	for column in df2:
		createSubtitleFile(sdf, df2[column], column, sheet)
