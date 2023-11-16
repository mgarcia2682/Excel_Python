import pandas as pd
import datetime
import os
import openpyxl
import re

# Put the special characters here
special_char = ['cpe:/o', '/', ':', '_', '.', '(', ')']
special_char_escaped = list(map(re.escape, special_char))

# Asks what you want to name the file
#file_name = input('What would you like to name the file ?  ')

modified_file_path = input('Where would you like to save the file to ? ')

# Asks where the file is located that you would like to use
excel_file_path = input('Where is the excel file located :  ')

# Asks what to filter the excel file by
column_filter = input('What column would you like to fiter by ?  ')

# Reads the excel file and imports it into a DataFrame
df = pd.read_excel(excel_file_path)

# Replaces the special characters from the list in special_char list above
df[column_filter] = df[column_filter].replace(special_char_escaped, ' ', regex=True)

# Slices the characters after the first 31 from the Plugin colum in the DataFrame
df[column_filter] = df[column_filter].str.slice(0,31)

# Writes the new DataFrame to a XLSX named from the "finame" variable using the xlsxwriter. Each unique value wil be added to its own sheet in the workbook.
writer = pd.ExcelWriter(modified_file_path,engine= 'xlsxwriter')
for value in df[column_filter].unique():
    newdf = df[df[column_filter] == value]
    newdf.to_excel(writer,sheet_name = value, index = False)

# Closes the workbook so it can save.
writer.close()