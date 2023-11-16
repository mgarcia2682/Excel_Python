import pandas as pd
import datetime
import os
import openpyxl
import re

# Put the special characters here
special_char = ['cpe:/o', '/', ':', '_', '.']
special_char_escaped = list(map(re.escape, special_char))

# Asks what you want to name the file
print("Enter the the name you would like to save the file as")
file_name = input()

# Path to file ony use "r" if using a folder path
excel_file_path = r"D:\<yourfilepathhere>.xlsx"

# Reads the excel file and imports it into a DataFrame
df =pd.read_excel(excel_file_path)
print(df)

df['<Column Name Here>'] = df['<Column Name Here>'].replace(special_char_escaped, ' ', regex=True)
print(df)

# Slices the characters after the first 30 from the Plugin colum in the DataFrame
df['<Column Name Here>'] = df['<Column Name Here>'].str.slice(0,31)

# Writes the new DataFrame to a XLSX named from the "finame" variable using the xlsxwriter. Each unique value wil be added to its own sheet in the workbook.
writer = pd.ExcelWriter(file_name + ".xlsx",engine= 'xlsxwriter')
for value in df['<Column Name Here>'].unique():
    newdf = df[df['<Column Name Here>'] == value]
    newdf.to_excel(writer,sheet_name = value, index = False)

# Closes the workbook so it can be saved.
writer.close()