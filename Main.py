import AlgemeneInfo
import Data
import Grafieken
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from os import listdir

class Main():
    huidigUur = input("Hoeveel uren zijn de zonnepanelen gestressed?")
    # --------------- wisselen tussen JW en NSP ---------------
    wb = openpyxl.load_workbook('__PID_BIFI_NPERT_JW_5BB.xlsx', data_only=True)
    generalSheet = wb['General']
    sheetNames = ['JW1_F', 'JW1_B', 'JW2_F', 'JW2_B']
    uren = AlgemeneInfo.AlgemeneInfo.datasheet(wb, sheetNames, huidigUur)

    #fonts
    vet = Font(size=14, bold=True)
    groot = Font(size=14)

    for n in range(0, len(sheetNames), 1):
        #manier zoeken om de range te beginnen vanaf het eerste nieuw toegevoegde uur
        for uur in range(0, len(uren), 1):
            activeSheet = wb[sheetNames[n]]

            nieuwPad = './' + str(uren[uur]) + '/' + sheetNames[n]

            ivPad = nieuwPad + '/IV/' + sheetNames[n]
            eqePad = nieuwPad + '/IQE/'
            drk = Data.Data.getDataList(str(ivPad) + '.drk')
            lgt = Data.Data.getDataList(str(ivPad) + '.lgt')

            vDark = drk[0]
            iDark = drk[1]
            vLight = lgt[0]
            iLight = lgt[1]

            #uur
            column1 = activeSheet.max_column + 2
            column2 = activeSheet.max_column + 5
            activeSheet.merge_cells(start_row = 1, start_column = column1, end_row = 1, end_column = column2)
            activeSheet.cell(row=1, column=column1).value = str(uren[uur]) + ' h'
            activeSheet.cell(row=1, column=column1).font = vet
            activeSheet.cell(row=1, column=column1).alignment = Alignment(horizontal='center')

            #light dark
            column3 = column1 + 1
            activeSheet.merge_cells(start_row = 2, start_column = column1, end_row = 2, end_column = column3)
            activeSheet.cell(row=2, column=column1).value = 'light'
            activeSheet.cell(row=2, column=column1).font = groot
            activeSheet.cell(row=2, column=column1).alignment = Alignment(horizontal='center')

            column4 = column1 + 2
            column5 = column4 + 1
            activeSheet.merge_cells(start_row = 2, start_column = column4, end_row = 2, end_column = column5)
            activeSheet.cell(row=2, column=column4).value = 'dark'
            activeSheet.cell(row=2, column=column4).font = groot
            activeSheet.cell(row=2, column=column4).alignment = Alignment(horizontal='center')

            #iv invullen
            activeSheet.cell(row=3, column=column1).value = 'V'
            activeSheet.cell(row=3, column=column3).value = 'I'
            activeSheet.cell(row=3, column=column4).value = 'V'
            activeSheet.cell(row=3, column=column5).value = 'I'

            #data invullen
            # --------------- verschil lengte light / dark ? ---------------
            for j in range(0,len(vLight),1):
                activeSheet.cell(row=j+4, column=column1).value = vLight[j]
                activeSheet.cell(row=j+4, column=column3).value = iLight[j]
            for j in range(0,len(vDark),1):
                activeSheet.cell(row=j+4, column=column4).value = vDark[j]
                activeSheet.cell(row=j+4, column=column5).value = iDark[j]

            # ------ EQE ------
            eqeSheet = 'EQE_' + sheetNames[n]
            activeSheet = wb[eqeSheet]
            #titels invullen
            column1 = activeSheet.max_column + 2
            column2 = activeSheet.max_column + 3
            activeSheet.merge_cells(start_row=1, start_column=column1, end_row=1, end_column=column2)
            activeSheet.cell(row=1, column=column1).value = str(uren[uur]) + ' h'

            activeSheet.cell(row=2, column=column1).value = 'EQE'
            activeSheet.cell(row=2, column=column2).value = 'EQE norm.'

            #data uit file halen
            eqeFile = ''
            for f in listdir(eqePad):
                if f.startswith(sheetNames[n]) and f.endswith('.eqe'):
                    eqeFile = f
            eqe = Data.Data.getDataList(str(eqePad) + eqeFile)[1]

            #data invullen
            for j in range(0,len(eqe),1):
                activeSheet.cell(row=j+3, column=column1).value = eqe[j]
                activeSheet.cell(row=j + 3, column=column2).value = eqe[j] / activeSheet.cell(row=j + 3, column=3).value

        #grafieken
        Grafieken.Grafieken.makeChart(sheetNames[n], wb, uren)
        Grafieken.Grafieken.makeChart(eqeSheet, wb, uren)
    Grafieken.Grafieken.makeSeperateGraphs(wb, sheetNames, uren)

    print('Saving...')
    wb.save('__PID_BIFI_NPERT_JW_5BB_updated.xlsx')


