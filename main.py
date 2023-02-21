from Utils import GetValidDatas, EscreverCabecalho, ConverterStrParaData, getSleepData, escreverSleepStages, escreverLevelSummary, GetSleepScore, GetStressDoDia, GetDatasetDoDia, GerarDataSetFinal, ConverterDeXLSXToCSV
import json
import csv
import xlsxwriter as xl
from datetime import timedelta


fSleep = open('Sleep/sleep-2023-01-17.json') # alterar nome do ficheiro
dataSleep = json.load(fSleep)

fileSleepScore = open("Sleep/sleep_score.csv", "r") # manter nome do ficheiro
dataSleepScore = list(csv.DictReader(fileSleepScore)) # os dados devem estar separados por , e não por ;
fileSleepScore.close()

fileStress = open("Stress.csv", "r") # csv do IATP2_Dataset_Contexto1_Dados
dataStress = list(csv.DictReader(fileStress)) # os dados devem estar separados por , e não por ;
fileStress.close()

fileEvery = open("Dataset.csv", "r") # csv do IATP2_Dataset_Contexto2_Dados
dataDataSet = list(csv.DictReader(fileEvery)) # os dados devem estar separados por , e não por ;
fileStress.close()

firstCalorias = [2319, 2248, 2525, 2649, 2604, 2240, 2944, 2219, 2428, 3017, 2350, 2444, 2397, 2452, 2379] # caloria gastas por dia, do 1 dia ao ultimo dia

firstPassos = [3400, 3287, 6463, 6874, 6765, 2856, 9996, 3081, 5048, 13814, 4299, 4905, 3417, 5074,4314] # passos dados por dia, do 1 dia ao ultimo dia

# firstHeartRate = [61, 63, 63, 65, 65, 66, 67, 65, 64, 62, 60, 62, 63, 63, 63] # media de batimentos cardiacos diarios, do dia 1 ao ultimo dia

firstDistance = [2.47, 2.39, 4.66, 4.96, 4.95, 2.07, 7.26, 2.24, 3.59, 10.04, 3.11, 3.56, 2.48, 3.68, 3.14] # distancia percorrida em km por cada dia, do 1 dia ao ultimo

firstFloors = [8, 2, 9, 7, 21, 1, 11, 0, 8, 22, 9, 5, 2, 17, 15] # andares percorridos em cada dia, do 1 dia ao ultimo dia

firstAge = 20 # idade

fistBMI = 24.1 # bmi

firstDatas = GetValidDatas(dataSleepScore)
dataSleep.reverse()

dayCounter = 0

print("A gerar ficheiros...")

while dayCounter < len(firstDatas):

    sleepData = getSleepData(dataSleep, firstDatas[dayCounter])
    tempList =sleepData['levels']['data']

    nome = firstDatas[dayCounter] + '.xlsx'
    workbook = xl.Workbook(nome)
    worksheet = workbook.add_worksheet()

    EscreverCabecalho(worksheet)

    #format1 = workbook.add_format({'num_format': 'hh:mm'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-ddThh:mm:ss'})
    format5 = workbook.add_format({'num_format': 'hh:mm:ss'})

    startSleep = sleepData['startTime']
    endSleep = sleepData['endTime']

    startSleepData = ConverterStrParaData(startSleep)
    endSleepData = ConverterStrParaData(endSleep)
    row = 1
    while startSleepData <= endSleepData:

        worksheet.write_number(row, 0, dayCounter+1) # dia
        worksheet.write_number(row, 1, row) # epoch
        worksheet.write(row, 2, startSleepData, format5)

        escreverSleepStages(tempList, worksheet, startSleepData, row)

        worksheet.write(row, 4, startSleepData.strftime("%d/%m/%Y"))

        worksheet.write_datetime(row, 5, startSleepData, format3)
        worksheet.write_datetime(row, 6, endSleepData, format3)

        worksheet.write_number(row, 7, sleepData['minutesToFallAsleep'])
        worksheet.write_number(row, 8, sleepData['minutesAsleep'])
        worksheet.write_number(row, 9, sleepData['minutesAwake'])
        worksheet.write_number(row, 10, sleepData['minutesAfterWakeup'])
        worksheet.write_number(row, 11, sleepData['timeInBed'])
        worksheet.write_number(row, 12, sleepData['efficiency'])

        levelsSummary = sleepData['levels']['summary']

        escreverLevelSummary(worksheet, levelsSummary, row)

        worksheet.write_number(row, 21, int(GetSleepScore(dataSleepScore, firstDatas[dayCounter])))

        worksheet.write_number(row, 22, firstCalorias[dayCounter])

        worksheet.write_number(row, 23, firstPassos[dayCounter])

        worksheet.write_number(row, 24, firstDistance[dayCounter])

        worksheet.write_number(row, 25, firstFloors[dayCounter])

        worksheet.write_number(row, 26, firstAge)

        worksheet.write(row, 27, fistBMI)

        stressDoDia = GetStressDoDia(dataStress, firstDatas[dayCounter])

        worksheet.write_number(row, 28, int(stressDoDia['Stress1']))
        worksheet.write_number(row, 29, int(stressDoDia['Stress2']))
        worksheet.write_number(row, 30, int(stressDoDia['Stress3']))

        datasetDoDia = GetDatasetDoDia(dataDataSet, firstDatas[dayCounter])

        worksheet.write_number(row, 31, float(datasetDoDia['Temperatura Media do ar']))
        worksheet.write_number(row, 32, int(datasetDoDia['Abertura de Apps (Total)']))
        worksheet.write_number(row, 33, int(datasetDoDia['Notificacoes (Total)']))
        worksheet.write_number(row, 34, int(datasetDoDia['Desbloqueios (Total)']))
        worksheet.write_number(row, 35, int(datasetDoDia['0-3am']))
        worksheet.write_number(row, 36, int(datasetDoDia['3-6am']))
        worksheet.write_number(row, 37, int(datasetDoDia['6-9am']))
        worksheet.write_number(row, 38, int(datasetDoDia['9am-12pm']))
        worksheet.write_number(row, 39, int(datasetDoDia['12-3pm']))
        worksheet.write_number(row, 40, int(datasetDoDia['3-6pm']))
        worksheet.write_number(row, 41, int(datasetDoDia['6-9pm']))
        worksheet.write_number(row, 42, int(datasetDoDia['9pm-0am']))
        worksheet.write_number(row, 43, int(datasetDoDia['Enviados (Total)']))
        worksheet.write_number(row, 44, int(datasetDoDia['Recebidos (Total)']))

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

