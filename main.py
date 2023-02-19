from Utils import GetValidDatas, EscreverCabecalho, ConverterStrParaData, getSleepData, escreverSleepStages, escreverLevelSummary, GetSleepScore, GetStressDoDia, GetDatasetDoDia, GerarDataSetFinal, ConverterDeXLSXToCSV
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

firstCalorias = [2319, 2248, 2525, 2649, 2604, 2240, 2944, 2219, 2428, 3017, 2350, 2444, 2397, 2452, 2379]

firstPassos = [3400, 3287, 6463, 6874, 6765, 2856, 9996, 3081, 5048, 13814, 4299, 4905, 3417, 5074,4314]

# firstHeartRate = [61, 63, 63, 65, 65, 66, 67, 65, 64, 62, 60, 62, 63, 63, 63]

firstDistance = [2.47, 2.39, 4.66, 4.96, 4.95, 2.07, 7.26, 2.24, 3.59, 10.04, 3.11, 3.56, 2.48, 3.68, 3.14]

firstFloors = [8, 2, 9, 7, 21, 1, 11, 0, 8, 22, 9, 5, 2, 17, 15]

firstAge = 20

fistBMI = 24.1

firstDatas = GetValidDatas(dataSleepScore)
dataSleep.reverse()
dataSleepScore.reverse()

dayCounter = 0

print("A gerar ficheiros...")

while dayCounter < len(firstDatas):

    sleepData = getSleepData(dataSleep, firstDatas[dayCounter])
    tempList =sleepData['levels']['data']

    nome = firstDatas[dayCounter] + '.xlsx'
    workbook = xl.Workbook(nome)
    worksheet = workbook.add_worksheet()

    EscreverCabecalho(worksheet)

    format1 = workbook.add_format({'num_format': 'hh:mm'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-ddThh:mm:ss'})
    format5 = workbook.add_format({'num_format': 'hh:mm:ss'})

    startSleep = sleepData['startTime']
    endSleep = sleepData['endTime']

    startSleepData = ConverterStrParaData(startSleep)
    endSleepData = ConverterStrParaData(endSleep)
    row = 1
    while startSleepData <= endSleepData:

        worksheet.write(row, 0, dayCounter+1) # dia
        worksheet.write(row, 1, row) # epoch
        worksheet.write(row, 2, startSleepData, format5)

        escreverSleepStages(tempList, worksheet, startSleepData, row)

        worksheet.write(row, 4, startSleepData.strftime("%d/%m/%Y"))

        worksheet.write_datetime(row, 5, startSleepData, format3)
        worksheet.write_datetime(row, 6, endSleepData, format3)

        worksheet.write(row, 7, sleepData['minutesToFallAsleep'])
        worksheet.write(row, 8, sleepData['minutesAsleep'])
        worksheet.write(row, 9, sleepData['minutesAwake'])
        worksheet.write(row, 10, sleepData['minutesAfterWakeup'])
        worksheet.write(row, 11, sleepData['timeInBed'])
        worksheet.write(row, 12, sleepData['efficiency'])

        levelsSummary = sleepData['levels']['summary']

        escreverLevelSummary(worksheet, levelsSummary, row)

        worksheet.write(row, 21, int(GetSleepScore(dataSleepScore, firstDatas[dayCounter])))

        worksheet.write(row, 22, firstCalorias[dayCounter])

        worksheet.write(row, 23, firstPassos[dayCounter])

        worksheet.write(row, 24, firstDistance[dayCounter])

        worksheet.write(row, 25, firstFloors[dayCounter])

        worksheet.write(row, 26, firstAge)

        worksheet.write(row, 27, fistBMI)

        stressDoDia = GetStressDoDia(dataStress, firstDatas[dayCounter])

        worksheet.write(row, 28, int(stressDoDia['Stress1']))
        worksheet.write(row, 29, int(stressDoDia['Stress2']))
        worksheet.write(row, 30, int(stressDoDia['Stress3']))

        datasetDoDia = GetDatasetDoDia(dataDataSet, firstDatas[dayCounter])

        worksheet.write(row, 31, int(datasetDoDia['Temperatura Media do ar']))
        worksheet.write(row, 32, int(datasetDoDia['Abertura de Apps (Total)']))
        worksheet.write(row, 33, int(datasetDoDia['Notificacoes (Total)']))
        worksheet.write(row, 34, int(datasetDoDia['Desbloqueios (Total)']))
        worksheet.write(row, 35, int(datasetDoDia['0-3am']))
        worksheet.write(row, 36, int(datasetDoDia['3-6am']))
        worksheet.write(row, 37, int(datasetDoDia['6-9am']))
        worksheet.write(row, 38, int(datasetDoDia['9am-12pm']))
        worksheet.write(row, 39, int(datasetDoDia['12-3pm']))
        worksheet.write(row, 40, int(datasetDoDia['3-6pm']))
        worksheet.write(row, 41, int(datasetDoDia['6-9pm']))
        worksheet.write(row, 42, int(datasetDoDia['9pm-0am']))
        worksheet.write(row, 43, int(datasetDoDia['Enviados (Total)']))
        worksheet.write(row, 44, int(datasetDoDia['Recebidos (Total)']))

        worksheet.write(row, 45, datasetDoDia['1 email enviado (timestamp)'])
        worksheet.write(row, 46, datasetDoDia['ultimo email enviado (timestamp)'])

        try:
            worksheet.write(row, 47, int(datasetDoDia['Total (efetuadas + recebidas)']))
        except:
            print()
        worksheet.write(row, 48, datasetDoDia['1 chamada (timestamp)'])
        worksheet.write(row, 49, datasetDoDia['ultima chamada (timestamp)'])
        try:
            worksheet.write(row, 50, int(datasetDoDia['Total (recebidas + enviadas)']))
            worksheet.write(row, 51, int(datasetDoDia['Total apos 0h (recebidas + enviadas)']))
        except:
            print()
        worksheet.write(row, 52, datasetDoDia['1 mensagem (timestamp)'])
        worksheet.write(row, 53, datasetDoDia['ultima mensagem (timestamp)'])

        row += 1
        startSleepData = startSleepData + timedelta(seconds=30)

    workbook.close()

    dayCounter += 1

print("Ficheiros gerados.")

print("A gerar dataset final...")
GerarDataSetFinal()
ConverterDeXLSXToCSV()
print("Dataset final gerado.")

