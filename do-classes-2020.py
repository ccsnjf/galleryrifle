
from typing import Any, Union
import xlsxwriter
import shutil
import csv
import operator
import argparse
import numpy as np
import pandas as pd

# Define paths
#input_path = ('C:\\Dropbox\\My Dropbox\\_classifications-spreadsheets\\')
input_path = ('C:\\Dropbox\\My Dropbox\\_stats\\2020\\')
output_path = ('C:\\Users\\ccsnjf\\PycharmProjects\\Classes\\')
google_path = ('C:\\_GDRIVE\\Google Drive\\Classifications\\')


parser = argparse.ArgumentParser(description = "In 2017 - 209	SAW  | 210	ATSC  | 211	JSPC-Spring  | 212	Basildon  | 213	Western Winner \
                                                214	Phoenix  | 215	WCSA Open  | 216	Wapinschaw  |217 Derby  | 218 Scottish Pistol Shorts \
                                                219	Welsh Open | 220	Scottish 1500 Open | 221 Nationals | 222 SLG Bisley \
                                                223	JSPC-Autumn | 224 CLSTSA Burnley | 225 AAW",
                                 epilog = "For 2020 - 269	Chelmsford Spring | 270	SAW | 271	ATSC | 272	JSPC-Spring | 273	Basildon \
                                            274	Mattersey | 275	Western Winner| 276	Phoenix |277 Southern Counties Open  |278 Wapinschaw \
                                            279	Derby | 280	Scottish Shorts | 281 Welsh Open | 282 Nationals  |283	SLG Bisley  |284 JSPC-Autumn \
                                            285	CLSTSA Burnley | 286 AAW  | 287	Chelmsford Winter | 288	FDPC Rimfire Festival")
#parser.add_argument('venue', type=int, default=False, help='VenueID (int) to start parsing from. 209 is SAW-2017. ')
parser.add_argument("-v", "--venue", required=False, type=int, help="Starting venue - must be an integer")
parser.add_argument("-e", "--end", required=False, type=int, help="Ending venue - must be an integer")
parser.add_argument("-c", "--copy", required=False, type=int, help="Set copy flag to 1 to publish")

args = parser.parse_args()
#print(args.venue)

'''
2017
209	SAW  | 210	ATSC  | 211	JSPC-Spring  | 212	Basildon  | 213	Western Winner 
214	Phoenix  | 215	WCSA Open  | 216	Wapinschaw  |217 Derby  | 218 Scottish Pistol Shorts 
219	Welsh Open | 220	Scottish 1500 Open | 221 Nationals | 222 SLG Bisley 
223	JSPC-Autumn | 224 CLSTSA Burnley | 225 AAW

2018
226	SAW |  227 JSPC-Spring  | 228	ATSC | 229	Basildon|  230	Western Winner 
231	WCSA Open|  232	Phoenix|  233	Wapinschaw | 234	Derby 2018| 235	Scottish Pistol Shorts 
236	Welsh Open | 237	Nationals | 238	Scottish 1500 Open | 239 SLG Bisley | 240 JSPC-Autumn 
241	CLSTSA Burnley | 242 AAW | 243 FDPC Rimfire Festival 

2019
250	SAW | 251 ATSC | 252	JSPC-Spring | 253 Basildon 
254	Western Winner | 255	Southern Counties Open | 256 Phoenix |257 Southern Counties Open | 258	Wapinschaw 
259	Derby | 260	Scottish Shorts Z 261	Welsh Open | 262 Nationals | 263 SLG Bisley 
264	JSPC-Autumn  | 265 CLSTSA Burnley | 266	AAW |  267  Chelmsford | 268 FDPC Rimfire Festival 

2020
269	Chelmsford Spring | 270	SAW | 271	ATSC | 272	JSPC-Spring | 273	Basildon 
274	Mattersey | 275	Western Winner| 276	Phoenix |277 Southern Counties Open  |278 Wapinschaw 
279	Derby | 280	Scottish Shorts | 281 Welsh Open | 282 Nationals  |283	SLG Bisley  |284 JSPC-Autumn 
285	CLSTSA Burnley | 286 AAW  | 287	Chelmsford Winter | 288	FDPC Rimfire Festival 
'''

#parse the master scores file and split out into individual events

venue_start = args.venue

if args.end is not None:
    venue_end = args.end
else:
    venue_end = 289
# Read in the raw Scores file here
# Contains all the scores for the last three+ years
# We need to filter out all scores but the classified stuff

# Define Venue IDs for the Class window
# venue ID to be defined as input arg - venue_start
# end element range the end of 2020 season - currently FDPC RFF if not defined

venueids = range(venue_start,venue_end)
#venueids = range(209,289)

inputfile = 'all-scores-last 3yrs.csv'

with open(input_path + inputfile, newline='') as readin:

    print('Raw scores data in from ' + input_path + inputfile + '\n')
    data = pd.read_csv(readin, encoding='utf8', index_col = False)
    print('Slurping ' + str(len(data)) + ' input scores from the main input file ')
    data = data[data['VenueID'].isin(venueids)]
    print('Processing ' + str(len(data)) + ' Scores for pre-classification filter run.')

    data = data[['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score', 'xcount']]

    #print(data.head())

readin.close()
data.to_csv('modernscores.csv', index=False)

#Filter out the Classified events - as defined in the events array below

events = [701, 702, 721, 722, 901, 902, 921, 922, 1101, 1102, 1121, 1122, \
          1301, 1302, 1321, 1322, 1501, 1502, 1521, 1522, 1601, 1602, 1701, 1702, 1721, 1722, 1901, 1902, 1921, 1922]

with open('modernscores.csv', newline='') as readin:
    data = pd.read_csv(readin, encoding='utf8', index_col = False)
    data = data[data['EventNo'].isin(events)]
    #print(data.head())
    print('Saved ' + str(len(data)) + ' Scores for classification processing.\n')
readin.close()
data.to_csv('the dogs bollocks-last 3 years.csv', index=False)


# do some tidying up
# combine score and xcount from the input DB - output to a local txt file

outfile = output_path + 'thedb.txt'
inputfile = 'the dogs bollocks-last 3 years.csv'
#inputfile = 'the dogs bollocks-last 5 years.csv'
#inputfile = 'the dogs bollocks.csv'

print("Now working on classified scores data from - " + inputfile)
print("Starting at eventNo " + str(venue_start))

with open(outfile, mode='w', newline='') as event_file:
    event_file_writer = csv.writer(event_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
    event_file_writer.writerow(['GRID', 'Name', 'VenueID', 'Venue', 'EventNo', 'Event', 'Score'])

    with open(inputfile, newline='') as readin:
        csv_reader = csv.reader(readin)
        sort = sorted(csv_reader, key=operator.itemgetter(0))

        for row in sort:
            if (row[4]).isdigit(): #row[4] is the event number (701, 1101 etc)
                #print (row[4])
                xnumber = int(row[7]) / 1000 #row7 is xcount
                x = str(xnumber)
                xcount = x.split(".")
                score = row[6] + "." + xcount[1]  #concatenate the score into xxx.yyy
                #print (row[0], score)
                event_file_writer.writerow([row[0], row[1], row[2], row[3], row[4], row[5], score])
readin.close()

#####################################################################################################
#Now work from the local tidied up file
# Read in and parse the the DB to obtain max scores

#set the output MS Excel file
xlwriter = pd.ExcelWriter('highscores.xlsx', engine='xlsxwriter')

with open('thedb.txt', newline='') as readin:
    data = pd.read_csv(readin, encoding='utf8')

    # print(data.head())

    table1 = pd.pivot_table(data,
                            values='Score',
                            index=['GRID', 'Name'],
                            columns=['EventNo'],
                            aggfunc=np.max,
                            fill_value=0)

    table1.to_excel(xlwriter, sheet_name='Highest Scores', startrow=1)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets['Highest Scores']

    #dump out to CSV as well
    table1.to_csv('highscores.csv', float_format="%.3f")

xlwriter.save()

print("Dumped.... highscores.csv")

#Do this brute force for the time being.
#Also want the individual files for reference.

###############################701##################################

with open ('classes701.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('highscores.csv', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp1sb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #701
            if 300 <= float(row[2]) <= 301:
                row.append('X')
                class_file_writer.writerow(row)
            elif 298.000 <= float(row[2]) <= 299.030:
                row.append('A')
                class_file_writer.writerow(row)
            elif 294.000 <= float(row[2]) <= 297.030:
                row.append('B')
                class_file_writer.writerow(row)
            elif 285.000 <= float(row[2]) <= 293.030:
                row.append('C')
                class_file_writer.writerow(row)
            elif 1 <= float(row[2]) <= 284.030:
                row.append('D')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################702##################################

with open ('classes702.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes701.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp1cf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #702
                if 300.027 <= float(row[3]) <= 301:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 300.023 <= float(row[3]) <= 300.026:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 300.000 <= float(row[3]) <= 300.022:
                    row.append('B')
                    class_file_writer.writerow(row)
                elif 297.000 <= float(row[3]) <= 299.030:
                    row.append('C')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[3]) <= 296.030:
                    row.append('D')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################721##################################

with open ('classes721.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes702.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp1lbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #721
                if 299.000 <= float(row[4]) <= 301:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 292.000 <= float(row[4]) <= 298.030:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[4]) <= 291.030:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################722##################################

with open ('classes722.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes721.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp1lbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
            #722
            if 299.000 <= float(row[5]) <= 301:
                row.append('X')
                class_file_writer.writerow(row)
            elif 292.000 <= float(row[5]) <= 298.030:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1 <= float(row[5]) <= 291.030:
                row.append('B')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################901##################################

with open ('classes901.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes722.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp2sb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
            #901
            if 587 <= float(row[6]) <= 601:
                row.append('X')
                class_file_writer.writerow(row)
            elif 567.000 <= float(row[6]) <= 586.060:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1 <= float(row[6]) <= 586.060:
                row.append('B')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################902##################################

with open ('classes902.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes901.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp2cf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #902
                if 595.000 <= float(row[7]) <= 601:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 587.000 <= float(row[7]) <= 594.060:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[7]) <= 586.060:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################921##################################

with open ('classes921.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes902.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp2lbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #921
                if 588.000 <= float(row[8]) <= 601:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 568.000 <= float(row[8]) <= 587.060:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[8]) <= 567.060:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################922##################################

with open ('classes922.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes921.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("tp2lbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #922
                if 572.000 <= float(row[9]) <= 601:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 540.000 <= float(row[9]) <= 571.060:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[9]) <= 539.060:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1101##################################

with open ('classes1101.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes922.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("mtsb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1101
            if 118.000 <= float(row[10]) <= 121:
                row.append('X')
                class_file_writer.writerow(row)
            elif 113.000 <= float(row[10]) <= 117.024:
                row.append('A')
                class_file_writer.writerow(row)
            elif 106.000 <= float(row[10]) <= 112.024:
                row.append('B')
                class_file_writer.writerow(row)
            elif 95.000 <= float(row[10]) <= 105.024:
                row.append('C')
                class_file_writer.writerow(row)
            elif 1 <= float(row[10]) <= 95.024:
                row.append('D')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1102##################################

with open ('classes1102.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1101.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("mtcf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1102
                if 120.000 <= float(row[11]) <= 121:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 118.000 <= float(row[11]) <= 119.024:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 114.000 <= float(row[11]) <= 117.024:
                    row.append('B')
                    class_file_writer.writerow(row)
                elif 106.000 <= float(row[11]) <= 113.024:
                    row.append('C')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[11]) <= 105.024:
                    row.append('D')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1121##################################

with open ('classes1121.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1102.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("mtlbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1121
                if 116.000 <= float(row[12]) <= 121:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 107.000 <= float(row[12]) <= 115.024:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[12]) <= 106.024:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1122##################################

with open ('classes1122.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1121.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("mtlbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1122
                if 113.000 <= float(row[13]) <= 121:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 106.000 <= float(row[13]) <= 112.024:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[13]) <= 105.024:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1301##################################

with open ('classes1301.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1122.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("phxasb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1301
            if 189.000 <= float(row[14]) <= 201:
                row.append('X')
                class_file_writer.writerow(row)
            elif 175.000 <= float(row[14]) <= 188.040:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1 <= float(row[14]) <= 174.040:
                row.append('B')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1302##################################

with open ('classes1302.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1301.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("phxacf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1302
                if 197.000 <= float(row[15]) <= 201:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 188.000 <= float(row[15]) <= 196.040:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[15]) <= 187.040:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1321##################################

with open ('classes1321.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1302.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("phxalbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1321
                if 190.000 <= float(row[16]) <= 201:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 177.000 <= float(row[16]) <= 189.040:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[16]) <= 176.040:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1322##################################

with open ('classes1322.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1321.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("phxalbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1322
                if 183.000 <= float(row[17]) <= 201:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 163.000 <= float(row[17]) <= 182.040:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[17]) <= 162.040:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1501##################################

with open ('classes1501.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1322.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1500sb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1501
            if 1481.000 <= float(row[18]) <= 1501:
                row.append('X')
                class_file_writer.writerow(row)
            elif 1471.000 <= float(row[18]) <= 1480.150:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1448.000 <= float(row[18]) <= 1470.150:
                row.append('B')
                class_file_writer.writerow(row)
            elif 1375.000 <= float(row[18]) <= 1447.150:
                row.append('C')
                class_file_writer.writerow(row)
            elif 1 <= float(row[18]) <= 1374.150:
                row.append('D')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1502##################################

with open ('classes1502.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1501.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1500cf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1502
                if 1494.000 <= float(row[19]) <= 1501:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1490.000 <= float(row[19]) <= 1493.150:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1480.000 <= float(row[19]) <= 1489.150:
                    row.append('B')
                    class_file_writer.writerow(row)
                elif 1456.000 <= float(row[19]) <= 1479.150:
                    row.append('C')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[19]) <= 1455.150:
                    row.append('D')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1521##################################

with open ('classes1521.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1502.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1500lbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1521
                if 1485.000 <= float(row[20]) <= 1501:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1450.000 <= float(row[20]) <= 1484.150:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[20]) <= 1449.150:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1522##################################

with open ('classes1522.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1521.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1500lbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1522
                if 1468.000 <= float(row[21]) <= 1501:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1400.000 <= float(row[21]) <= 1467.150:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[21]) <= 1399.150:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1601##################################

with open ('classes1601.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1522.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1020sb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1501
            if 1016.000 <= float(row[22]) <= 1021:
                row.append('X')
                class_file_writer.writerow(row)
            elif 1003.000 <= float(row[22]) <= 1015.120:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1 <= float(row[22]) <= 1002.120:
                row.append('B')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1602##################################

with open ('classes1602.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1601.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("the1020cf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1602
                if 1020.000 <= float(row[23]) <= 1021:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1015.000 <= float(row[23]) <= 1019.120:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[23]) <= 1014.120:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1701##################################

with open ('classes1701.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1602.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("bianchisb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1701
            if 1836.000 <= float(row[24]) <= 1921:
                row.append('X')
                class_file_writer.writerow(row)
            elif 1701.000 <= float(row[24]) <= 1835.192:
                row.append('A')
                class_file_writer.writerow(row)
            elif 1 <= float(row[24]) <= 1700.192:
                row.append('B')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1702##################################

with open ('classes1702.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1701.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("bianchicf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1702
                if 1890.000 <= float(row[25]) <= 1921:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1801.000 <= float(row[25]) <= 1889.192:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[25]) <= 1800.192:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1721##################################

with open ('classes1721.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1702.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("bianchilbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1721
                if 1895.000 <= float(row[26]) <= 1921:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1750.000 <= float(row[26]) <= 1894.192:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[26]) <= 1749.192:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1722##################################

with open ('classes1722.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1721.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("bianchilbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1722
                if 1825.000 <= float(row[27]) <= 1921:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 1750.000 <= float(row[27]) <= 1824.192:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[27]) <= 1749.192:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1901##################################

with open ('classes1901.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1722.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("advtsb")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
        #1901
            if 290.000 <= float(row[28]) <= 301:
                row.append('X')
                class_file_writer.writerow(row)
            elif 281.000 <= float(row[28]) <= 289.060:
                row.append('A')
                class_file_writer.writerow(row)
            elif 272.000 <= float(row[28]) <= 280.060:
                row.append('B')
                class_file_writer.writerow(row)
            elif 250.000 <= float(row[28]) <= 271.060:
                row.append('C')
                class_file_writer.writerow(row)
            elif 1 <= float(row[28]) <= 249.060:
                row.append('D')
                class_file_writer.writerow(row)
            else:
                row.append('U')
                class_file_writer.writerow(row)
    readin.close()
class_writer.close()

###############################1902##################################

with open ('classes1902.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1901.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("advtcf")

        # class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1902
                if 180.000 <= float(row[29]) <= 181:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 178.000 <= float(row[29]) <= 179.036:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 174.000 <= float(row[29]) <= 177.0036:
                    row.append('B')
                    class_file_writer.writerow(row)
                elif 165.000 <= float(row[29]) <= 173.036:
                    row.append('C')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[29]) <= 164.036:
                    row.append('D')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()


###############################1921##################################

with open ('classes1921.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1902.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("advtlbp")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1921
                if 177.000 <= float(row[30]) <= 181:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 169.000 <= float(row[30]) <= 176.036:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[30]) <= 168.036:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

###############################1922##################################

with open ('classes1922.txt', 'w', newline = '', encoding="utf-8-sig") as class_writer:
    with open('classes1921.txt', newline = '', encoding="utf-8-sig") as readin:
        csv_reader = csv.reader(readin)
        class_file_writer = csv.writer(class_writer, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)
        headers = next(csv_reader, None)
        #next (csv_reader, None)
        headers.append("advtlbr")

        #class_file_writer(headers)
        class_file_writer.writerow(headers)

        for row in csv_reader:
                #1922
                if 174.000 <= float(row[31]) <= 181:
                    row.append('X')
                    class_file_writer.writerow(row)
                elif 163.000 <= float(row[31]) <= 173.036:
                    row.append('A')
                    class_file_writer.writerow(row)
                elif 1 <= float(row[31]) <= 162.036:
                    row.append('B')
                    class_file_writer.writerow(row)
                else:
                    row.append('U')
                    class_file_writer.writerow(row)

    readin.close()
class_writer.close()

#make a copy for the various output options required

shutil.copy2('classes1922.txt', 'scores_and_classes.csv')

print("Dumped.... scores_and_classes.csv")


xlwriter = pd.ExcelWriter('classes.xlsx', engine='xlsxwriter')
with open('scores_and_classes.csv', newline='', encoding="utf-8-sig") as readin:
    df = pd.read_csv(readin, encoding='utf8')

    # print(data.head())
    # print(data.tail())
    # print(data.info())

    # slice the DF. 0:2 == first 2 cols (GRID, Name) -30 == last 30 cols (the class numbers)
    # use ILOC - .iloc[<row selection>, <column selection>]
    new_df = pd.concat([df.iloc[:, 0:2], df.iloc[:, -30:]], axis=1)

    #print(new_df.head())

    new_df.to_excel(xlwriter, sheet_name='Classes', startrow=1, index=False)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets['Classes']

    #dump out to CSV as well
    new_df.to_csv('classes.csv', float_format="%.3f", index=True)

xlwriter.save()

print("Dumped.... classes.csv")

##################Combine into a single file - Classes and Scores###################################
# We got 'Scores and Classes' file - slice it to get 'Classes and Scores' file
# and we need a column 0 index for this for this to upload it to the GR website.

xlwriter = pd.ExcelWriter('classes_and_scores.xlsx', engine='xlsxwriter')
with open('scores_and_classes.csv', newline='', encoding="utf-8-sig") as readin:
    df = pd.read_csv(readin, encoding='utf8')


    # slice the DF. 0:2 == first 2 cols (GRID, Name) -30 == last 30 cols (the class numbers). 2:32 are the scores.
    # use ILOC - .iloc[<row selection>, <column selection>]
    new_df = pd.concat([df.iloc[:, 0:2], df.iloc[:, -30:], df.iloc[:, 2:32]], axis=1)
    new_df.index += 1

    new_df.to_excel(xlwriter, sheet_name='Sheet1', startrow=1, index=True)
    workbook = xlwriter.book
    worksheet = xlwriter.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '####0.000'})
    worksheet.set_column('AG:BJ', 7, format1)

    #dump out to CSV as well
    new_df.to_csv('classes_and_scores.csv', float_format="%.3f", index=True)

xlwriter.save()

#Copy the files where needed
copy_flag = 0
copy_flag = args.copy

if copy_flag == 1:
    # shutil.copy2('scores_and_classes.csv', google_path + 'scores_and_classes.csv')
    shutil.copy2('classes_and_scores.csv', google_path + 'classes_and_scores.csv')
    shutil.copy2('classes.csv', google_path + 'classes.csv')
    shutil.copy2('highscores.csv', google_path + 'highscores.csv')
    print("Copied files to google drive")
else:
    print ("**Not published** - did you remember the copy flag (-c1)?")

