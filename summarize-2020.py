from typing import Any, Union
import xlsxwriter
# from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
import shutil
import csv
import operator
import numpy as np
import pandas as pd
import os

google_path = ('C:\\_GDRIVE\\Google Drive\\RANKINGS\\')

summarydata = pd.read_csv("summary-rankings.txt", encoding='ansi')

# print(summary-data.head())
# print(summary-data.tail())
# print(summary-data.info())

#reading from summary-rankings.txt file
#Format is GRID,Name,No.Shot,Rank,Event


#concatenate     'strings' Rank and Shot - separate with /
summarydata['Ratio'] = summarydata.Rank.astype(str).str.cat(summarydata.Shot.astype(str), sep='/')

tablesumm = pd.pivot_table(summarydata,
                        values='Ratio',
                        index=['GRID', 'Name'],
                        columns=['Event'],
                        aggfunc=np.min)
#print (tablesumm.head())


xlwriter = pd.ExcelWriter('summary-rankings.xlsx', engine='xlsxwriter')

tablesumm.to_excel(xlwriter, sheet_name="Ranks", startrow=1)
workbook = xlwriter.book
worksheet = xlwriter.sheets["Ranks"]

format2 = workbook.add_format({'align': 'left'})
format3 = workbook.add_format({'bold': False, 'align': 'centre', 'font_color': 'blue', 'font_size': 12})
boldtitle = workbook.add_format({'bold': True, 'font_color': 'red', 'font_size': 24})
merge_format = workbook.add_format({'align': 'left'})

#num_blue = workbook.add_format({'font_color': 'blue'})
#num_black = workbook.add_format({'font_color': 'black'})

worksheet.set_column('B:B', 20)
worksheet.set_column('C:R', 8, format3)


worksheet.freeze_panes(2, 2)

# sort the title out - map number to event name via the dict above - print out both
worksheet.merge_range('A1:M1', "Summary", merge_format)

title = "Per Event Positions    Ranking Tables 2019: Position/Events Shot"
worksheet.write('A1', title, boldtitle)

xlwriter.save()

shutil.copy2('summary-rankings.xlsx', google_path + 'summary-rankings.xlsx')
