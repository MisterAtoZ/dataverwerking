import AlgemeneInfo
import Data
import Grafieken
import WorkbookLayout
import openpyxl
from os import listdir

class Main():
    def begin(self, pv, huidigUur, wbName, pad):
        sheetNames = [pv+'1_F',pv+'1_B',pv+'2_F',pv+'2_B']
        graphNames = ['%PID', 'FF', 'Voc', 'Isc']

        wb = openpyxl.load_workbook(pad + wbName, data_only=True)
        dataEx = openpyxl.load_workbook(pad + 'data-exchange' + huidigUur + '.xlsx')
        WorkbookLayout.WorkbookLayout.makeSheets(wb, dataEx, graphNames, sheetNames)
        info = AlgemeneInfo.AlgemeneInfo.datasheet(wb, sheetNames, dataEx)
        begin = info[0]
        uren = info[1]

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
        return True