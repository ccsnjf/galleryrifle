from typing import Any, Union
import xlsxwriter
import csv
import operator
import numpy as np
import pandas as pd

#parse the master scores file and split out into individual events

events = [701, 702, 721, 722, 1101, 1102, 1121, 1122,1501, 1502, 1521, 1522]

readin = open(r'C:\Dropbox\My Dropbox\_classifications-spreadsheets\rankings.csv', 'r')

for event in events:
    the_path = ('C:\\Users\\ccsnjf\\PycharmProjects\\testing\\')
    in_filename = str(event) + "." + 'txt'
    outfile = the_path + in_filename
    #print(outfile)
    with open(outfile, mode='w', newline='') as event_file:
        event_file_writer = csv.writer(event_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        event_file_writer.writerow(['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score'])

        with open(r'C:\Users\ccsnjf\PycharmProjects\testing\thedb.csv', newline='') as readin:
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

xlwriter = pd.ExcelWriter('zzz.xlsx', engine='xlsxwriter')

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
                      #margins=True,
                      #margins_name='Totals')
    #print (table1)
    #sum_column = table1.sum(axis=1)
    #print(sum_column)

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

    #move 'best2 to the beginning
    table1 = table1[ [ 'Best4' ] + [ col for col in table1.columns if col != 'Best4' ] ]

    #add order column
    table1.insert(0, 'Rank', range(1, 1+len(table1)))


    #table1 = table1.apply(best_four, axis=0)
    #print(table1.head())

    table1.to_csv('xxx.txt', float_format="%.3f")

    ##do export to MS Excel format here.
    table1.to_excel(xlwriter, sheet_name=event)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets[event]
    cell_format = workbook.add_format()
    cell_format.set_align('left')

    format1 = workbook.add_format({'num_format': '####0.000', 'align': 'left'})
    format2 = workbook.add_format({'align': 'left'})
    format3 = workbook.add_format({'align': 'centre'})
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 8, format3)
    worksheet.set_column('D:D', 10, format1)

xlwriter.save()