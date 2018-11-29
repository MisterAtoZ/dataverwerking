import AlgemeneInfo
import Data
import Grafieken
import WorkbookLayout
import openpyxl
import os
from os import listdir

class Main():
    def begin(self, pv, maxUur, aantal, kanten, wbName, pad):
        sheetNames = []
        for n in range(1,int(aantal)+1,1):
            for i in range(0,len(kanten),1):
                sheetNames.append(pv+str(n)+'_'+str(kanten[i]))
        print(str(sheetNames))
        graphNames = ['%PID', 'FF', 'Voc', 'Isc']

        subfolders = [f.name for f in os.scandir(pad) if f.is_dir()]
        subfolders.sort(key=int)
        wb = openpyxl.load_workbook(pad + wbName, data_only=True)
        if os.path.exists(pad + subfolders[len(subfolders)-1]):
            for f in listdir(pad + subfolders[len(subfolders)-1]):
                if f.startswith('data-exchange') and f.endswith('.xlsx'):
                    data_exchange = f
                    dataEx = openpyxl.load_workbook(pad + subfolders[len(subfolders)-1] + '/' + data_exchange)
                    WorkbookLayout.WorkbookLayout.makeSheets(wb, dataEx, graphNames, sheetNames)
                    info = AlgemeneInfo.AlgemeneInfo.datasheet(wb, sheetNames, dataEx, maxUur, subfolders)
                    begin = info[0]
                    uren = info[1]
        else:
            print('Error')
            return False

        for n in range(0, len(sheetNames), 1):
            #begin = eerste rij die ingevult moet worden (dus begin-1), maar de uren beginnen pas op rij 2 (dus nog eens -1) => begin-2
            for uur in range(begin-2, len(uren), 1):
                activeSheet = wb[sheetNames[n]]

                nieuwPad = pad + str(subfolders[uur]) + '/' + sheetNames[n]

                ivPad = nieuwPad + '/IV/'
                eqePad = nieuwPad + '/IQE/'

                #data uit file halen
                if os.path.exists(ivPad):
                    if os.path.exists(ivPad + sheetNames[n] + '.drk'):
                        drk = Data.Data.getDataList(str(ivPad) + sheetNames[n] + '.drk')
                    else:
                        print(str(ivPad) + '.drk file bestaat niet')
                    if os.path.exists(ivPad + sheetNames[n] + '.lgt'):
                        lgt = Data.Data.getDataList(str(ivPad) + sheetNames[n] + '.lgt')
                    else:
                        print(str(ivPad) + sheetNames[n] + ' .lgt file bestaat niet')
                else:
                    print(str(ivPad) + ' bestaat niet')

                WorkbookLayout.WorkbookLayout.setIV(activeSheet, uren[uur], drk, lgt)

                # ------ EQE ------
                eqeSheet = 'EQE_' + sheetNames[n]
                activeSheet = wb[eqeSheet]

                #data uit file halen
                eqeFile = ''
                if os.path.exists(eqePad):
                    for f in listdir(eqePad):
                        if f.startswith(sheetNames[n]) and f.endswith('.eqe'):
                            eqeFile = f
                    if eqeFile != '':
                        eqe = Data.Data.getDataList(str(eqePad) + eqeFile)[1]
                        WorkbookLayout.WorkbookLayout.setEQE(activeSheet, uren[uur], eqe)
                    else:
                        print(str(eqePad) + ' file bestaat niet')
                else:
                    print(str(eqePad) + ' bestaat niet')

            #grafieken
            Grafieken.Grafieken.makeChart(sheetNames[n], wb, uren)
            Grafieken.Grafieken.makeChart('EQE_' + sheetNames[n], wb, uren)
        Grafieken.Grafieken.makeSeperateGraphs(wb, graphNames, sheetNames, uren)

        print('Saving...')
        wb.save(pad + str(uren[len(uren)-1]) + wbName)
        return True