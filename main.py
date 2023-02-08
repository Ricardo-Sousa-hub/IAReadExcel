import json
import csv
import xlsxwriter as xl
from datetime import datetime
from datetime import timedelta


fSleep = open('Sleep/sleep-2023-01-17.json')
dataSleep = json.load(fSleep)

fileSleepScore = open("Sleep/sleep_score.csv", "r")
dataSleepScore = list(csv.DictReader(fileSleepScore))
fileSleepScore.close()

fileStress = open("Stress.csv", "r")
dataStress = list(csv.DictReader(fileStress))
fileStress.close()

fileEvery = open("Dataset.csv", "r")
dataDataSet = list(csv.DictReader(fileEvery))
fileStress.close()



row = 1
column = 0

firstDatas = ["2023-01-19", "2023-01-20", "2023-01-21", "2023-01-22", "2023-01-23", "2023-01-24", "2023-01-25",
              "2023-01-26", "2023-01-27", "2023-01-28", "2023-01-29", "2023-01-30", "2023-01-31", "2023-02-01",
              "2023-02-02"]

dias = len(firstDatas)
sleepStages = {
    'wake':0,
    'deep':1,
    'light':2,
    'rem':3
}

for dia in range(dias):

    nome = dataSleep[dia]['dateOfSleep']+'.xlsx'
    workbook = xl.Workbook(nome)
    worksheet = workbook.add_worksheet()

    worksheet.write(0,0,'TimeStamp')
    worksheet.write(0, 1, 'Sleep_Stage')
    worksheet.write(0, 4, 'Column1.levels.data.dateTime')
    worksheet.write(0, 5, 'timestamp')
    worksheet.write(0, 6, 'sleepstage')
    worksheet.write(0, 7, 'Column1.levels.data.level')


    lastDate = 'a'
    firstData2 = ""
    for i in dataSleep:

        levels = i['levels']
        listData = levels['data']
        for y in listData:
            if y['dateTime'][0:10] == firstDatas[dia]:



                worksheet.write(row, 4, y['dateTime'])
                worksheet.write(row, 5, y['dateTime'][10:18])
                data = y['dateTime'][10:18]
                if y['level'] == 'wake':
                    worksheet.write(row, 6, sleepStages['wake'])
                    #worksheet.write(row, 1, sleepStages['wake'])
                if y['level'] == 'deep':
                    worksheet.write(row, 6, sleepStages['deep'])
                    #worksheet.write(row, 1, sleepStages['deep'])
                if y['level'] == 'light':
                    worksheet.write(row, 6, sleepStages['light'])
                    #worksheet.write(row, 1, sleepStages['light'])
                if y['level'] == 'rem':
                    worksheet.write(row, 6, sleepStages['rem'])
                    #worksheet.write(row, 1, sleepStages['rem'])

                lastDate = y['dateTime']
                worksheet.write(row, 7, y['level'])
                row += 1

                if firstData2 == '':
                    firstData2 = y['dateTime']

                row2 = 1
                while firstData2 != lastDate:
                    worksheet.write(row2, 0, firstData2)
                    row2 += 1

                    d1 = firstData2.split('-')[2][0:2]
                    time = datetime(int(firstData2[0:4]), int(firstData2[5:7]), int(d1), int(firstData2[11:13]), int(firstData2[14:16]), int(firstData2[17:19]))
                    time2 = time + timedelta(seconds=30)

                    firstData2 = time2.strftime("%Y-%m-%dT%H:%M:%S.000")
        firstData2 = ""
        row = 1
    workbook.close()
