import AlgemeneInfo
import Data
import Grafieken
import WorkbookLayout
import openpyxl
from openpyxl.styles import Font
from os import listdir

class Main():
    # fonts
    vet = Font(size=14, bold=True)
    groot = Font(size=14)

    pv = input('welke zonnecellen zijn er getest geweest? (JW/NSP)')

    if pv == 'JW':
        wbName = '__PID_BIFI_npert_JW_5BB'
        sheetNames = ['JW1_F', 'JW1_B', 'JW2_F', 'JW2_B']
        pad = './20180910_BIFI_npert_JolyW_5BB/'
    elif pv == 'NSP':
        wbName = '__PID_BIFI_pperc_NSP_4BB'
        sheetNames = ['NSP1_F', 'NSP1_B', 'NSP2_F', 'NSP2_B']
        pad = './20180910_BIFI_pperc_NSP_4BB/'
    else:
        raise Exception('Geef een geldig antwoord. Uw input was: {}'.format(pv))

    huidigUur = input('Hoeveel uren zijn de zonnepanelen gestressed?')

    graphNames = ['%PID', 'FF', 'Voc', 'Isc']

    wb = openpyxl.load_workbook(pad + wbName + '.xlsx', data_only=True)
    dataEx = openpyxl.load_workbook(pad + 'data-exchange' + huidigUur + '.xlsx')
    WorkbookLayout.WorkbookLayout.makeSheets(wb, dataEx, graphNames, sheetNames)
    info = AlgemeneInfo.AlgemeneInfo.datasheet(wb, sheetNames, dataEx)
    begin = info[0]
    uren = info[1]
    generalSheet = wb['General']

    for n in range(0, len(sheetNames), 1):
        #begin = eerste rij die ingevult moet worden (dus begin-1), maar de uren beginnen pas op rij 2 (dus nog eens -1) => begin-2
        for uur in range(begin-2, len(uren), 1):
            activeSheet = wb[sheetNames[n]]

            nieuwPad = pad + str(uren[uur]) + '/' + sheetNames[n]

            ivPad = nieuwPad + '/IV/' + sheetNames[n]
            eqePad = nieuwPad + '/IQE/'

            #data uit file halen
            drk = Data.Data.getDataList(str(ivPad) + '.drk')
            lgt = Data.Data.getDataList(str(ivPad) + '.lgt')

            WorkbookLayout.WorkbookLayout.setIV(activeSheet, uren[uur], drk, lgt)

            # ------ EQE ------
            eqeSheet = 'EQE_' + sheetNames[n]
            activeSheet = wb[eqeSheet]

            #data uit file halen
            eqeFile = ''
            for f in listdir(eqePad):
                if f.startswith(sheetNames[n]) and f.endswith('.eqe'):
                    eqeFile = f
            eqe = Data.Data.getDataList(str(eqePad) + eqeFile)[1]

            WorkbookLayout.WorkbookLayout.setEQE(activeSheet, uur, eqe)

        #grafieken
        Grafieken.Grafieken.makeChart(sheetNames[n], wb, uren)
        Grafieken.Grafieken.makeChart('EQE_' + sheetNames[n], wb, uren)
    Grafieken.Grafieken.makeSeperateGraphs(wb, graphNames, sheetNames, uren)

    print('Saving...')
    wb.save(pad + wbName + '_updated.xlsx')


