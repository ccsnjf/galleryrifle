from typing import Any, Union
import xlsxwriter
import shutil
import csv
import operator
import argparse
import numpy as np
import pandas as pd
import os

try:
    os.remove("summary-rankings.txt")
except:
    print("Error while deleting file ", "summary-rankings.txt")

parser = argparse.ArgumentParser(description = "For 2019 - 250	SAW | 251 ATSC | 252	JSPC-Spring | 253 Basildon \
                                                254	Western Winner | 255	Southern Counties Open | 256 Phoenix |257 Southern Counties Open | 258	Wapinschaw \
                                                259	Derby | 260	Scottish Shorts Z 261	Welsh Open | 262 Nationals | 263 SLG Bisley \
                                                264	JSPC-Autumn  | 265 CLSTSA Burnley | 266	AAW |  267  Chelmsford | 268 FDPC Rimfire Festival",
                                 epilog = "For 2020 - 269	Chelmsford Spring | 270	SAW | 271	ATSC | 272	JSPC-Spring | 273	Basildon \
                                            274	Mattersey | 275	Western Winner| 276	Phoenix |277 Southern Counties Open  |278 Wapinschaw \
                                            279	Derby | 280	Scottish Shorts | 281 Welsh Open | 282 Nationals  |283	SLG Bisley  |284 JSPC-Autumn \
                                            285	CLSTSA Burnley | 286 AAW  | 287	Chelmsford Winter | 288	FDPC Rimfire Festival")
#parser.add_argument('venue', type=int, default=False, help='VenueID (int) to start parsing from. 209 is SAW-2017. ')
parser.add_argument("-v", "--venue", required=False, type=int, help="Starting venue - must be an integer")
parser.add_argument("-e", "--end", required=False, type=int, help="Ending venue - must be an integer")

args = parser.parse_args()

venue_start = args.venue
venue_end = args.end + 1

# define number event name mappings


comp_names = {
    101: "25m Precision GRSB",
    102: "25m Precision GRCF",
    103: "25m Precision GRCF Open",
    104: "25m Precision GRCF Classic",
    121: "25m Precision LPB",
    122: "25m Precision LBR",
    301: "50m Precision GRSB",
    302: "50m Precision GRCF",
    303: "50m Precision GRCF Open",
    304: "50m Precision GRCF Classic",
    321: "50m Precision LPB",
    322: "50m Precision LBR",
    501: "America Match GRSB",
    502: "America Match GRCF",
    503: "America Match GRCF Open",
    504: "America Match GRCF Classic",
    521: "America Match LPB",
    522: "America Match LBR",
    701: "T&P1 GRSB",
    702: "T&P1 GRCF",
    703: "T&P1 GRCF Open",
    704: "T&P1 GRCF Classic",
    721: "T&P1 LPB",
    722: "T&P1 LBR",
    724: "T&P1 LBP",
    725: "T&P1 LBR",
    901: "T&P2 GRSB",
    902: "T&P2 GRCF",
    903: "T&P2 GRCF Open",
    904: "T&P2 GRCF Classic",
    921: "T&P2 LPB",
    922: "T&P2 LBR",
    1021: "T&P3 LBP",
    1022: "T&P3 LBR",
    1101: "Multi-Target GRSB",
    1102: "Multi-Target GRCF",
    1103: "Multi-Target GRCF Open",
    1104: "Multi-Target GRCF Classic",
    1121: "Multi-Target LPB",
    1122: "Multi-Target LBR",
    1124: "Multi-Target LBP",
    1125: "Multi-Target LBR",
    1301: "Phoenix A GRSB",
    1302: "Phoenix A GRCF",
    1303: "Phoenix A GRCF Classic",
    1304: "Phoenix A GRCF Classic",
    1321: "Phoenix A LPB",
    1322: "Phoenix A LBR",
    1501: "1500 GRSB",
    1502: "1500 GRCF",
    1503: "1500 GRCF Open",
    1504: "1500 GRCF Classic",
    1521: "1500 LPB",
    1522: "1500 LBR",
    1524: "1500 LPB",
    1525: "1500 LBR",
    1601: "1020 GRSB",
    1602: "1020 GRCF",
    1603: "1020 GRCF Open",
    1604: "1020 GRCF Classic",
    1621: "1020 LPB",
    1622: "1020 LBR",
    1821: "WA48 LBP",
    1822: "WA48 LBR",
    1901: "Advancing Target GRSB",
    1902: "Advancing Target GRCF",
    1903: "Advancing Target GRCF Classic",
    1904: "Advancing Target Classic",
    1921: "Advancing Target LPB",
    1922: "Advancing Target LBR",
    1924: "Advancing Target LPB",
    1925: "Advancing Target LBR",
    2621: "NRA Rapids LBP",
    2622: "NRA Rapids LBR",
}

# Define paths
#input_path = ('C:\\Dropbox\\My Dropbox\\_classifications-spreadsheets\\')
input_path = ('C:\\Dropbox\\My Dropbox\\_stats\\2020\\')
#input_path = ('C:\\Users\\ccsnjf\\PycharmProjects\\Classes\\')
output_path = ('C:\\Users\\ccsnjf\\PycharmProjects\\rankings\\rankings-2020\\')
google_path = ('C:\\_GDRIVE\\Google Drive\\RANKINGS\\')

# parse the master scores file and split out into individual events
print ("Reading from " + input_path + " ...done")

venueids = range(venue_start, venue_end)
#venueids = range(241,269)

print ("parsing from " + str(venue_start) + " to " + str(venue_end))
inputfile = 'all-scores-last 3yrs.csv'

with open(input_path + inputfile, newline='') as readin:

    print('Raw scores data in from ' + input_path + inputfile + '\n')
    data = pd.read_csv(readin, encoding='utf8', index_col = False)
    print('Slurping ' + str(len(data)) + ' input scores from the main input file')
    data = data[data['VenueID'].isin(venueids)]
    print('Processing ' + str(len(data)) + ' Scores for pre-rankings filter run.')

    data = data[['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score', 'xcount']]

    #print(data.head())

readin.close()
data.to_csv('the dogs bollocks-last 3 years.csv', index=False)


events = [701, 702, 721, 722, 1101, 1102, 1121, 1122, 1501, 1502, 1521, 1522, 1601, 1602, 1621, 1622]

for event in events:
    in_filename = str(event) + "." + 'txt'
    outfile = output_path + in_filename
    # print(outfile)
    with open(outfile, mode='w', newline='') as event_file:
        event_file_writer = csv.writer(event_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        event_file_writer.writerow(['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score'])

        with open('the dogs bollocks-last 3 years.csv', newline='') as readin:
        #with open(input_path + 'rankings.csv', newline='') as readin:
            csv_reader = csv.reader(readin)
            sort = sorted(csv_reader, key=operator.itemgetter(0))
            for row in sort:
                if (row[4]).isdigit():  # row[4] is the event number (701, 1101 etc)
                    if int(row[4]) == event:  # event = arg passed into this function
                        xnumber = int(row[7]) / 1000  # row7 is xcount
                        x = str(xnumber)
                        xcount = x.split(".")
                        score = row[6] + "." + xcount[1]  # concatenate the score into xxx.yyy
                        event_file_writer.writerow([row[0], row[1], row[2], row[3], row[4], row[5], score])

readin.close()

# Read in and parse the separate event files
print ("Writing to " + output_path + " ...done")
events = ["701", "702", "721", "722", "1101", "1102", "1121", "1122", "1501", "1502", "1521", "1522", "1601", "1602", "1621", "1622"]

xlwriter = pd.ExcelWriter('rankings.xlsx', engine='xlsxwriter')

for event in events:
    event_file = str(event) + "." + 'txt'
    data = pd.read_csv(event_file, encoding='ansi')

    # print(data.head())
    # print(data.tail())
    # print(data.info())

    table1 = pd.pivot_table(data,
                            values='Score',
                            index=['GRID', 'Name'],
                            columns=['VenueID'],
                            aggfunc=np.sum,
                            fill_value=0)

    # manual mapping of event IDs to display names for a season
    # need to keep these updated on a per season basis
    table1.rename(columns={241: 'LANCS18',
                           242: 'AAW18',
                           243: 'FDPC-RFF18',
                           250: 'SAW19',
                           251: 'ATSC19',
                           252: 'JSP-S19',
                           253: 'BAS19',
                           254: 'WW19',
                           256: 'PHO19',
                           257: 'SCO19',
                           258: 'WAP19',
                           259: 'DER19',
                           260: 'SCT19',
                           261: 'WLSH19',
                           262: 'NAT19',
                           263: 'SLG19',
                           264: 'JSP-A19',
                           265: 'LANCS19',
                           266: 'AAW19',
                           267: 'CRCA19',
                           268: 'RFF19',
                           269: 'CRCS20',
                           270: 'SAW20',
                           271: 'ATSC20',
                           272: 'JSP-S20',
                           273: 'BAS20',
                           274: 'MAT20',
                           275: 'WW20',
                           276: 'PHO20',
                           277: 'SCO20',
                           278: 'WAP20',
                           279: 'DER20',
                           280: 'SCT20',
                           281: 'WEL20',
                           282: 'NAT20',
                           283: 'SLG20',
                           284: 'JSP-A20',
                           285: 'LANCS20',
                           286: 'AAW20',
                           287: 'CRCA20',
                           288: 'RFF20'
                           },
                  inplace=True)



    # obtain top 4 scores by sifting into a separate table
    table2 = table1.stack().groupby(level=0).nlargest(4).unstack().reset_index(level=1, drop=True).reindex(
        columns=table1.columns)

    # sum the top 4 scores in table2 - append to table1
    table1['Best4'] = table2.apply(np.sum, axis=1)

    # sort on best4 column
    table1 = table1.sort_values('Best4', ascending=False)

    # move 'best4 to to the beginning
    table1 = table1[['Best4'] + [col for col in table1.columns if col != 'Best4']]

    # add order column
    table1.insert(0, 'Rank', range(1, 1 + len(table1)))

    # Strip out zeros - replace with whitespace (probably should not have added them in the pivot DF!)
    table1 = table1._get_numeric_data()
    table1[table1 <= 0] = ""

    # insert number of events shot column
    # non zero count  minus 2 to discount the rank and the 'best4' columns
    table1.insert(1, 'Shot', (table1.astype(bool).sum(axis=1)-2))

    #print(table1.head())

    table1.to_csv(event + '.csv', float_format="%.3f")

    #print (event)

##############  Keep a copy of summary info to build the summary rankings with  #####################################
    with open('summary-rankings.txt', 'a',newline='') as f:
        summary = table1.iloc[:, 0:2].copy()
        # take first column from table 1 - all the rows - 2 columns - [GRID NAME], Shot and Position
        #print (summary)
        # Add new column - "the event"
        summary.loc[:, 'Event'] = event
        if int(event) == 701:
            summary.to_csv(f, float_format="%.3f", header=True)
        else:
            summary.to_csv(f, float_format="%.3f", header=False)
#####################################################################################################################

    ##do export to MS Excel format here.
    ##Various bits of beautifying required
    table1.to_excel(xlwriter, sheet_name=event, startrow=1)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets[event]

    # The problem with doing zebra shading here with XLSXWriter is the format preference is
    # Cell -> Row -> Column. Rows overwrite column formatting even if added 'later'.
    # Left out for the time being.
    # row_shading1 = workbook.add_format({'align': 'centre', 'bg_color': '#ffcccc'})
    # row_shading2 = workbook.add_format({'align': 'centre', 'bg_color': '#cce5ff'})
    # for row in range(2, row_count + 2, 2):
    #    worksheet.set_row(row, cell_format=row_shading1)
    #    worksheet.set_row(row + 1, cell_format=row_shading2)

    # Format some columns here - Add a header (title)
    format1 = workbook.add_format({'num_format': '####0.000', 'align': 'centre', 'italic': True, 'bold': True, })
    format2 = workbook.add_format({'num_format': '####0', 'align': 'centre', 'italic': True })
    format3 = workbook.add_format({'bold': True, 'align': 'centre', 'font_color': 'red', 'font_size': 12})
    format4 = workbook.add_format({'num_format': '####0.000'})
    boldtitle = workbook.add_format({'bold': True, 'font_color': 'red', 'font_size': 24})
    merge_format = workbook.add_format({'align': 'left'})

    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 8, format3)
    worksheet.set_column('D:D', 6, format2)
    worksheet.set_column('E:E', 10, format1)
    worksheet.set_column('F:W', 7, format4)

    worksheet.freeze_panes(2, 2)

    # sort the title out - map number to event name via the dict above - print out both
    worksheet.merge_range('A1:J1', event, merge_format)
    name = comp_names[int(event)]
    title = "Event " + str(event) + " - " + name + " -  Ranking Tables 2019"
    print(name + " ...done")
    worksheet.write('A1', title, boldtitle)

xlwriter.save()

# copy over to the Gdoc sync file
print ("Copy to Google - " + google_path + " ...done")
shutil.copy2('rankings.xlsx', google_path + 'rankings.xlsx')

"""
##################################################################################################################
#Summarise the rankings on a per event basis here
#dump out to a separate spreadsheet - copy over to the Gdrive

summarydata = pd.read_csv("summary-rankings.txt", encoding='ansi')

# print(summary-data.head())
# print(summary-data.tail())
# print(summary-data.info())

#reading from summary-rankings.txt file
#Format is GRID,Name,Rank,Event
tablesumm = pd.pivot_table(summarydata,
                        values='Rank',
                        index=['GRID', 'Name'],
                        columns=['Event'],
                        aggfunc=np.min)
#print (tablesumm.head())


xlwriter = pd.ExcelWriter('summary-rankings.xlsx', engine='xlsxwriter')

tablesumm.to_excel(xlwriter, sheet_name="Ranks", startrow=1)
workbook = xlwriter.book
worksheet = xlwriter.sheets["Ranks"]

format2 = workbook.add_format({'align': 'left'})
format3 = workbook.add_format({'bold': True, 'align': 'centre', 'font_color': 'blue', 'font_size': 12})
boldtitle = workbook.add_format({'bold': True, 'font_color': 'red', 'font_size': 24})
merge_format = workbook.add_format({'align': 'left'})

worksheet.set_column('B:B', 20)
worksheet.set_column('C:N', 8, format3)

worksheet.freeze_panes(2, 2)

# sort the title out - map number to event name via the dict above - print out both
worksheet.merge_range('A1:J1', "Summary", merge_format)

title = "Per Event Positions    Ranking Tables 2019"
worksheet.write('A1', title, boldtitle)

xlwriter.save()

#shutil.copy2('summary-rankings.xlsx', google_path + 'summary-rankings.xlsx')
"""
