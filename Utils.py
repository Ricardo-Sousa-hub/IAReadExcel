import os
from collections import deque
from datetime import datetime
import glob
import pandas as pd
import openpyxl


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


def GetValidDatas(sleepScore):
    lista = []
    for item in sleepScore:
        itemT = item['timestamp'].split('T')
        itemT[0] = itemT[0].replace('-', ' ')
        itemT[0] = reverseText(itemT[0])
        itemT[0] = itemT[0].replace(' ', '-')
        lista.append(itemT[0])

    lista.reverse()

    return lista


def EscreverCabecalho(worksheet):
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
    worksheet.write(0, 22, "CaloriesBurned")
    worksheet.write(0, 23, "Steps")
    worksheet.write(0, 24, "Distance")
    worksheet.write(0, 25, "Floors")
    worksheet.write(0, 26, "Age")
    worksheet.write(0, 27, "BMI")
    worksheet.write(0, 28, 'stress_1')
    worksheet.write(0, 29, 'stress_2')
    worksheet.write(0, 30, 'stress_3')
    worksheet.write(0, 31, 'Abertura de Apps (Total)')
    worksheet.write(0, 32, 'Notificacoes (Total)')
    worksheet.write(0, 33, 'Desbloqueios (Total)')
    worksheet.write(0, 34, 'Tempo de Tela (minutos)')



def ConverterStrParaData(data):
    data = data.split('T')
    dataD = data[0].split('-')
    dataH = data[1].split(':')
    dataS = dataH[2].split('.')
    return datetime(year=int(dataD[0]), month=int(dataD[1]), day=int(dataD[2]), hour=int(dataH[0]),
                    minute=int(dataH[1]), second=int(dataS[0]))


def getSleepData(sleepData, data):
    for item in sleepData:
        dataTest = item['dateOfSleep']
        dataTest = dataTest.replace('-', ' ')
        dataTest = reverseText(dataTest)
        dataTest = dataTest.replace(' ', '-')
        if dataTest == data:
            return item


def escreverSleepStages(tempList, startTimeD):
    status = "0"
    for i in tempList:

        dateTime1 = ConverterStrParaData(i['dateTime'])

        if startTimeD == dateTime1:
            if i['level'] == 'wake':
                status = "0"
            if i['level'] == 'light':
                status = "2"
            if i['level'] == 'deep':
                status = "1"
            if i['level'] == 'rem':
                status = "3"
            return status
        else:
            print()


def escreverLevelSummary(worksheet, levelsSummary, count):
    try:
        worksheet.write_number(count, 13, levelsSummary['deep']['count'])
        worksheet.write_number(count, 14, levelsSummary['deep']['minutes'])
    except:
        print()
    try:
        worksheet.write_number(count, 15, levelsSummary['wake']['count'])
        worksheet.write_number(count, 16, levelsSummary['wake']['minutes'])
    except:
        print()
    try:
        worksheet.write_number(count, 17, levelsSummary['light']['count'])
        worksheet.write_number(count, 18, levelsSummary['light']['minutes'])
    except:
        print()
    try:
        worksheet.write_number(count, 19, levelsSummary['rem']['count'])
        worksheet.write_number(count, 20, levelsSummary['rem']['minutes'])
    except:
        print()


def GetSleepScore(dataSleepScore, data):
    for item in dataSleepScore:
        dataTest = item['timestamp']
        dataTest = dataTest.split('T')
        dataTest[0] = dataTest[0].replace('-', ' ')
        dataTest[0] = reverseText(dataTest[0])
        dataTest[0] = dataTest[0].replace(' ', '-')
        if dataTest[0] == data:
            return item['overall_score']


def GetStressDoDia(dataStress, data):
    for item in dataStress:
        print(item)
        dataTest = item['Data']
        dataTest = dataTest.replace('/', '-')
        if dataTest == data:
            return item


def GetDatasetDoDia(datasetDoDia, data):
    print(datasetDoDia)
    for item in datasetDoDia:
        dataTest = item['Data']
        dataTest = dataTest.replace('/', '-')
        if dataTest == data:
            return item


def GerarDataSetFinal():
    file_list = glob.glob("*.xlsx")
    file_list = sorted(file_list, key=lambda t: os.stat(t).st_mtime) # organizar ficheiros pela ordem pela qual foram gerados
    excl_list = []

    for file in file_list:
        excl_list.append(pd.read_excel(file))

    excl_merged = pd.DataFrame()

    for excl_file in excl_list:
        # appends the data into the excl_merged
        # dataframe.
        excl_merged = excl_merged.append(
            excl_file, ignore_index=True)

    excl_merged.to_excel('Final_Dataset.xlsx', index=False)


def ConverterDeXLSXToCSV():
    read_file = pd.read_excel("Final_Dataset.xlsx")
    read_file.to_csv('Final_Dataset.csv', index=None, header=True)
