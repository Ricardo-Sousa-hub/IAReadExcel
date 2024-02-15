from Utils import GetValidDatas, EscreverCabecalho, ConverterStrParaData, getSleepData, escreverSleepStages, escreverLevelSummary, GetSleepScore, GetStressDoDia, GetDatasetDoDia, GerarDataSetFinal, ConverterDeXLSXToCSV
import json
import csv
import xlsxwriter as xl
from datetime import timedelta


fSleep = open('Sleep/sleep-2024-01-11.json') # alterar nome do ficheiro
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

firstCalorias = [1821, 2444, 2582, 2851, 2547, 1859, 1821, 1983, 1925] # caloria gastas por dia, do 1 dia ao ultimo dia

firstPassos = [5105, 2629, 3289, 2178, 1675, 179, 1244, 1843, 1415] # passos dados por dia, do 1 dia ao ultimo dia

# firstHeartRate = [61, 63, 63, 65, 65, 66, 67, 65, 64, 62, 60, 62, 63, 63, 63] # media de batimentos cardiacos diarios, do dia 1 ao ultimo dia

firstDistance = [3.13, 1.91, 1.42, 1.56, 1.2, 0.13, 0.89, 1.32, 1.01] # distancia percorrida em km por cada dia, do 1 dia ao ultimo

firstFloors = [0, 8, 7, 6, 8, 1, 0, 0, 0] # andares percorridos em cada dia, do 1 dia ao ultimo dia

firstAge = 22 # idade

fistBMI = 24.7 # bmi

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
    print(tempList)
    print(startSleepData)
    status1 = int(escreverSleepStages(tempList, startSleepData))

    while startSleepData <= endSleepData:

        worksheet.write_number(row, 0, dayCounter+1) # dia
        worksheet.write_number(row, 1, row) # epoch
        worksheet.write(row, 2, startSleepData, format5)

        status = escreverSleepStages(tempList, startSleepData)

        if status != "null":
            try:
                status1 = int(status)
            except BaseException:
                print()

        worksheet.write_number(row, 3, status1)

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

        worksheet.write_number(row, 31, int(datasetDoDia['Abertura de Apps']))
        worksheet.write_number(row, 32, int(datasetDoDia['Notificacoes']))
        worksheet.write_number(row, 33, int(datasetDoDia['Desbloqueios do dispositivo']))
        worksheet.write_number(row, 34, int(datasetDoDia['Tempo de Tela (minutos)']))

        row += 1
        startSleepData = startSleepData + timedelta(seconds=30)

    workbook.close()

    dayCounter += 1

print("Ficheiros gerados.")

print("A gerar dataset final...")
GerarDataSetFinal()
ConverterDeXLSXToCSV()
print("Dataset final gerado.")

