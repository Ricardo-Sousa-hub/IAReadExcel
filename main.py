import json
import csv
import xlsxwriter as xl
from datetime import datetime
from datetime import timedelta
from collections import deque
import glob
import pandas as pd
import openpyxl

listaExcel = []


def reverseText(s):
    # base case
    if not s:
        return s

    # `s[low…high]` forms a word
    low = high = 0

    # create an empty stack
    stack = deque()

    # scan the text
    for i, c in enumerate(s):
        # if space is found, we found a word
        if c == ' ':
            # push each word into the stack
            stack.append(s[low:high + 1])

            # reset `low` and `high` for the next word
            low = high = i + 1
        else:
            high = i

    # push the last word into the stack
    stack.append(s[low:])

    # construct the string by following the LIFO order
    sb = ""
    while stack:
        sb += stack.pop() + ' '

    return sb[:-1]  # remove last space


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

# firstCalorias = []

# firstPassos = []

# firstHeartRate = []

# firstDistance = []

# firstFloors = []

# firstAge =

# fistBMI = KG/H

dias = len(firstDatas)
sleepStages = {
    'wake': 0,
    'deep': 1,
    'light': 2,
    'rem': 3
}

print("Generating excel...")
print("0%", end="")

index = 0
for dia in firstDatas:
    tempList = []
    dataS = dataSleep[index]['dateOfSleep']

    while dataS != dia:
        dataS = dataSleep[index]['dateOfSleep']
        index += 1

    tempList = dataSleep[index]['levels']['data']
    nome = dataSleep[index]['dateOfSleep'] + '.xlsx'

    workbook = xl.Workbook(nome)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Dia_Sessao')
    worksheet.write(0, 1, 'Epoch_ID')
    worksheet.write(0, 2, 'TimeStamp')
    worksheet.write(0, 3, 'Sleep_Stage')
    worksheet.write(0, 4, 'dateOfSleep')
    worksheet.write(0, 5, 'startTime')
    worksheet.write(0, 6, 'endTime')
    worksheet.write(0, 7, 'minutesToFallAsleep')
    worksheet.write(0, 8, 'minutesAsleep')
    worksheet.write(0, 9, 'minutesAwake')
    worksheet.write(0, 10, 'minutesAfterWakeup')
    worksheet.write(0, 11, 'timeInBed')
    worksheet.write(0, 12, 'efficiency')
    worksheet.write(0, 13, 'deep.count')
    worksheet.write(0, 14, 'deep.minutes')
    worksheet.write(0, 15, 'wake.count')
    worksheet.write(0, 16, 'wake.minutes')
    worksheet.write(0, 17, 'light.count')
    worksheet.write(0, 18, 'light.minutes')
    worksheet.write(0, 19, 'rem.count')
    worksheet.write(0, 20, 'rem.minutes')
    worksheet.write(0, 21, 'SleepScore')
    worksheet.write(0, 22, 'stress_1')
    worksheet.write(0, 23, 'stress_2')
    worksheet.write(0, 24, 'stress_3')
    worksheet.write(0, 25, 'Temp_Med')
    worksheet.write(0, 26, 'Abertura de Apps (Total)')
    worksheet.write(0, 27, 'Notificacoes (Total)')
    worksheet.write(0, 28, 'Desbloqueios (Total)')
    worksheet.write(0, 29, 'Tela 0-3am')
    worksheet.write(0, 30, 'Tela 3-6am')
    worksheet.write(0, 31, 'Tela 6-9am')
    worksheet.write(0, 32, 'Tela 9am-12pm')
    worksheet.write(0, 33, 'Tela 12-3pm')
    worksheet.write(0, 34, 'Tela 3-6pm')
    worksheet.write(0, 35, 'Tela 6-9pm')
    worksheet.write(0, 36, 'Tela 9pm-0am')
    worksheet.write(0, 37, 'Email1')
    worksheet.write(0, 38, 'Email2')
    worksheet.write(0, 39, 'Email3')
    worksheet.write(0, 40, 'Email4')
    worksheet.write(0, 41, 'call1')
    worksheet.write(0, 42, 'call2')
    worksheet.write(0, 43, 'call3')
    worksheet.write(0, 44, 'Text1')
    worksheet.write(0, 45, 'Text2')
    worksheet.write(0, 46, 'Text3')
    worksheet.write(0, 47, 'Text4')

    dia = firstDatas.index(dia) + 1
    worksheet.write(1, 0, dia)

    startSleep = ""
    endSleep = ""
    sleepScore = 0
    for data in dataSleep:
        if data['dateOfSleep'] == dataS:
            startSleep = data['startTime']
            endSleep = data['endTime']

    for score in dataSleepScore:
        new = score['timestamp'].split('T')
        if new[0] == dataS:
            sleepScore = score['overall_score']

    stressDoDia = {}

    for stress in dataStress:
        new = stress['Data'].replace('/', ' ')
        new = reverseText(new)
        new = new.replace(' ', '-')
        if new == dataS:
            stressDoDia = stress

    datasetDoDia = {}

    for dataset in dataDataSet:
        new = dataset['Dia'].replace('/', ' ')
        new = reverseText(new)
        new = new.replace(' ', '-')
        if new == dataS:
            datasetDoDia = dataset

    startTime = startSleep.split('T')[1].split(':')
    endTime = endSleep.split('T')[1].split(':')

    startDate = startSleep.split('T')[0].split('-')
    endDate = endSleep.split('T')[0].split('-')

    startTimeD = datetime(year=int(startDate[0]), month=int(startDate[1]), day=int(startDate[2]),
                          hour=int(startTime[0]), minute=int(startTime[1]), second=int(startTime[2].split('.')[0]))
    endTimeD = datetime(year=int(endDate[0]), month=int(endDate[1]), day=int(endDate[2]), hour=int(endTime[0]),
                        minute=int(endTime[1]), second=int(endTime[2].split('.')[0]))

    count = 1
    format1 = workbook.add_format({'num_format': 'hh:mm'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-ddThh:mm:ss'})
    format5 = workbook.add_format({'num_format': 'hh:mm:ss'})

    temp = dataSleep[index]['levels']['summary']

    try:
        if temp['deep']['count'] != 0 and temp['wake']['count'] != 0 and temp['light']['count'] != 0 and temp['rem'][
            'count'] != 0:
            while startTimeD < endTimeD:
                worksheet.write(count, 0, dia)
                worksheet.write(count, 1, count)
                worksheet.write_datetime(count, 2, startTimeD, format5)
                startTimeD = startTimeD + timedelta(seconds=30)

                for i in tempList:
                    time1 = i['dateTime'].split('T')
                    time2 = time1[0].split('-')
                    time3 = time1[1].split(':')

                    dateTime1 = datetime(year=int(time2[0]), month=int(time2[1]), day=int(time2[2]),
                                         hour=int(time3[0]), minute=int(time3[1]), second=int(time3[2].split('.')[0]))

                    if startTimeD < dateTime1:
                        status = 0
                        if i['level'] == 'wake':
                            status = 0
                        if i['level'] == 'light':
                            status = 2
                        if i['level'] == 'deep':
                            status = 1
                        if i['level'] == 'rem':
                            status = 3
                        worksheet.write(count, 3, status)
                        break

                worksheet.write(count, 4, startTimeD.strftime("%d/%m/%Y"))
                startTimeD1 = datetime(year=int(startDate[0]), month=int(startDate[1]), day=int(startDate[2]),
                                       hour=int(startTime[0]), minute=int(startTime[1]),
                                       second=int(startTime[2].split('.')[0]))
                worksheet.write_datetime(count, 5, startTimeD1, format3)
                endTimeD1 = datetime(year=int(endDate[0]), month=int(endDate[1]), day=int(endDate[2]),
                                     hour=int(endTime[0]), minute=int(endTime[1]), second=int(endTime[2].split('.')[0]))
                worksheet.write_datetime(count, 6, endTimeD1, format3)

                worksheet.write(count, 7, dataSleep[index]['minutesToFallAsleep'])
                worksheet.write(count, 8, dataSleep[index]['minutesAsleep'])
                worksheet.write(count, 9, dataSleep[index]['minutesAwake'])
                worksheet.write(count, 10, dataSleep[index]['minutesAfterWakeup'])
                worksheet.write(count, 11, dataSleep[index]['timeInBed'])
                worksheet.write(count, 12, dataSleep[index]['efficiency'])

                levelsSummary = dataSleep[index]['levels']['summary']
                try:
                    worksheet.write(count, 13, levelsSummary['deep']['count'])
                    worksheet.write(count, 14, levelsSummary['deep']['minutes'])
                except:
                    print()
                try:
                    worksheet.write(count, 15, levelsSummary['wake']['count'])
                    worksheet.write(count, 16, levelsSummary['wake']['minutes'])
                except:
                    print()
                try:
                    worksheet.write(count, 17, levelsSummary['light']['count'])
                    worksheet.write(count, 18, levelsSummary['light']['minutes'])
                except:
                    print()
                try:
                    worksheet.write(count, 19, levelsSummary['rem']['count'])
                    worksheet.write(count, 20, levelsSummary['rem']['minutes'])
                except:
                    print()
                worksheet.write(count, 21, int(sleepScore))

                worksheet.write(count, 22, int(stressDoDia['Stress1']))
                worksheet.write(count, 23, int(stressDoDia['Stress2']))
                worksheet.write(count, 24, int(stressDoDia['Stress3']))

                worksheet.write(count, 25, int(datasetDoDia['Temperatura Media do ar']))
                worksheet.write(count, 26, int(datasetDoDia['Abertura de Apps (Total)']))
                worksheet.write(count, 27, int(datasetDoDia['Notificacoes (Total)']))
                worksheet.write(count, 28, int(datasetDoDia['Desbloqueios (Total)']))
                worksheet.write(count, 29, int(datasetDoDia['0-3am']))
                worksheet.write(count, 30, int(datasetDoDia['3-6am']))
                worksheet.write(count, 31, int(datasetDoDia['6-9am']))
                worksheet.write(count, 32, int(datasetDoDia['9am-12pm']))
                worksheet.write(count, 33, int(datasetDoDia['12-3pm']))
                worksheet.write(count, 34, int(datasetDoDia['3-6pm']))
                worksheet.write(count, 35, int(datasetDoDia['6-9pm']))
                worksheet.write(count, 36, int(datasetDoDia['9pm-0am']))
                worksheet.write(count, 37, int(datasetDoDia['Enviados (Total)']))
                worksheet.write(count, 38, int(datasetDoDia['Recebidos (Total)']))

                worksheet.write(count, 39, datasetDoDia['1 email enviado (timestamp)'])
                worksheet.write(count, 40, datasetDoDia['ultimo email enviado (timestamp)'])
                try:
                    worksheet.write(count, 41, int(datasetDoDia['Total (efetuadas + recebidas)']))
                except:
                    print()
                worksheet.write(count, 42, datasetDoDia['1 chamada (timestamp)'])
                worksheet.write(count, 43, datasetDoDia['ultima chamada (timestamp)'])
                try:
                    worksheet.write(count, 44, int(datasetDoDia['Total (recebidas + enviadas)']))
                    worksheet.write(count, 45, int(datasetDoDia['Total apos 0h (recebidas + enviadas)']))
                except:
                    print()
                worksheet.write(count, 46, datasetDoDia['1 mensagem (timestamp)'])
                worksheet.write(count, 47, datasetDoDia['ultima mensagem (timestamp)'])

                count += 1
            index = 0

            workbook.close()
    except:
        index = 0
        print()

    print("=", end="")

print("100%")

print("Agrupando e organizando ficheiros resultantes da operação anterior...")
print("0%", end="")
print("=", end="")

file_list = glob.glob("*.xlsx")

excl_list = []

for file in file_list:
    excl_list.append(pd.read_excel(file))
    print("=", end="")

print("100%")
print("Gerando Dataset Final...")
print("0%", end="")

# create a new dataframe to store the
# merged excel file.
excl_merged = pd.DataFrame()

for excl_file in excl_list:
    # appends the data into the excl_merged
    # dataframe.
    excl_merged = excl_merged.append(
        excl_file, ignore_index=True)
    print("=", end="")
# exports the dataframe into excel file with
# specified name.
excl_merged.to_excel('Final_Dataset.xlsx', index=False)

print("100%")

print("Programa terminado")
