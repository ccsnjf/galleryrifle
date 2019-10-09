from typing import Any, Union
import xlsxwriter
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
import shutil
import csv
import operator
import numpy as np
import pandas as pd

#parse the master scores file and split out into individual events

events = [701, 702, 721, 722, 1101, 1102, 1121, 1122,1501, 1502, 1521, 1522]

#readin = open(r'C:\Dropbox\My Dropbox\_classifications-spreadsheets\rankings.csv', 'r')

for event in events:
    the_path = ('C:\\Users\\ccsnjf\\PycharmProjects\\testing\\')
    in_filename = str(event) + "." + 'txt'
    outfile = the_path + in_filename
    #print(outfile)
    with open(outfile, mode='w', newline='') as event_file:
        event_file_writer = csv.writer(event_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        event_file_writer.writerow(['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score'])

        with open(r'C:\Dropbox\My Dropbox\_classifications-spreadsheets\rankings.csv', newline='') as readin:
            csv_reader = csv.reader(readin)
            sort = sorted(csv_reader, key=operator.itemgetter(0))
            for row in sort:
                if (row[4]).isdigit(): #row[4] is the event number (701, 1101 etc)
                    if int(row[4]) == event: #event = arg passed into this function
                        xnumber = int(row[7]) / 1000 #row7 is xcount
                        x = str(xnumber)
                        xcount = x.split(".")
                        score = row[6] + "." + xcount[1]  #concatenate the score into xxx.yyy
                        event_file_writer.writerow([row[0], row[1], row[2], row[3], row[4], row[5], score])


readin.close()


#Read in and parse the separate event files

events = ["701", "702", "721", "722", "1101", "1102", "1121", "1122","1501", "1502", "1521", "1522"]

xlwriter = pd.ExcelWriter('rankings.xlsx', engine='xlsxwriter')

for event in events:
    event_file = str(event) + "." + 'txt'
    data = pd.read_csv(event_file, encoding='ansi')

    #print(data.head())
    #print(data.tail())
    #print(data.info())

    table1=pd.pivot_table(data,
                      values='Score',
                      index=['GRID','Name'],
                      columns=['VenueID'],
                      aggfunc=np.sum,
                      fill_value=0)


    #manual mapping of event IDs to display names for a season
    table1.rename(columns={241 :'LANCS18',
                        242:'AAW18',
                        243:'FDPC-RFF18',
                        250:'SAW19',
                        251:'ATSC19',
                        252:'JSP-S19',
                        253:'BAS19',
                        254:'WW19',
                        256:'PHO19',
                        257:'SCO19',
                        258:'WAP19',
                        259:'DER19',
                        260:'SCT19',
                        261:'WLSH19',
                        262:'Nat19',
                        263:'SLG19',
                        264:'JSP-A19',
                        265:'LANCS19',
                        266:'AAW19',
                        267:'CRC19',
                        268:'FDPC-RFF19'},
                 inplace=True)

    print (table1.head())

    #obtain top 4 scores by sifting into a separate table
    table2 = table1.stack().groupby(level=0).nlargest(4).unstack().reset_index(level=1, drop=True).reindex(columns=table1.columns)

    #sum the top 4 scores in table2 - append to table1
    table1['Best4'] = table2.apply(np.sum, axis=1)

    #sort on best4 column
    table1 = table1.sort_values('Best4', ascending=False)

    #move 'best4 to to the beginning
    table1 = table1[ [ 'Best4' ] + [ col for col in table1.columns if col != 'Best4' ] ]

    #add order column
    table1.insert(0, 'Rank', range(1, 1+len(table1)))
    #Strip out zeros - replace with whitespace
    table1 = table1._get_numeric_data()
    table1[table1 <= 0] = ""

    #print(table1.head())
    #table1.to_csv('xxx.txt', float_format="%.3f")

    row_count = table1.shape[0]  # row count
    col_count = table1.shape[1]  # col count

    ##do export to MS Excel format here.
    ##Various bits of beautifying required
    table1.to_excel(xlwriter, sheet_name=event, startrow=1)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets[event]

    # The problem with doing zebra shading here with XLSXWriter is the format preference is
    # Cell -> Row -> Column. Rows overwrite column formatting even if added 'later'.
    # Left out for the time being.
    #row_shading1 = workbook.add_format({'align': 'centre', 'bg_color': '#ffcccc'})
    #row_shading2 = workbook.add_format({'align': 'centre', 'bg_color': '#cce5ff'})
    #for row in range(2, row_count + 2, 2):
    #    worksheet.set_row(row, cell_format=row_shading1)
    #    worksheet.set_row(row + 1, cell_format=row_shading2)


    #Format some columns here - Add a header (title)
    format1 = workbook.add_format({'num_format': '####0.000', 'align': 'left'})
    format2 = workbook.add_format({'align': 'left'})
    format3 = workbook.add_format({'align': 'centre'})
    bold = workbook.add_format({'bold': True, 'font_color': 'red', 'font_size': 24})
    merge_format = workbook.add_format({'align': 'left'})

    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 8, format3)
    worksheet.set_column('D:D', 10, format1)


    worksheet.merge_range('A1:G1', event, merge_format)
    title = "Event " + str(event) + " Ranking Tables 2018-2019"
    worksheet.write('A1', title, bold)

xlwriter.save()

#copy over to the Gdoc sync file
shutil.copy2('rankings.xlsx', 'C:\\_GDRIVE\\Google Drive\\RANKINGS\\rankings.xlsx')
