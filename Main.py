import AlgemeneInfo
import Data
import Grafieken
import WorkbookLayout
import Photo
import openpyxl
import os
from os import listdir

class Main():


    # filePath = 'C:\\Users\Melissa\Documents\\01 Masterproef imec\Engineering\Python script\PID PSC-20190220T080521Z-001\PID PSC\\'
    # wbName='PID_PSC.xlsx'
    # wb = openpyxl.load_workbook(filePath + wbName, data_only=True)
    # files = [f.name for f in os.scandir(filePath)]
    # fileNames = []
    # print(files)
    # for f in files:
    #     if f.endswith('.dat'):
    #         fileNames.append(f[:-4])
    # print(str(fileNames))
    # fileNamesSorted = WorkbookLayout.WorkbookLayout.natural_sort(fileNames)
    # print(str(fileNamesSorted))
    # graphNames = ['%PID']
    # sheetNames = []
    # sampleNames = []
    # times = []
    # sampleNumbers = []
    # sampleMeasures = []
    # for f in fileNamesSorted:
    #     fileNamesSplitted = f.split('-')
    #     sampleNames.append(fileNamesSplitted[0])
    #     times.append(int(fileNamesSplitted[1][:-3]))
    #     if len(fileNamesSplitted)==4:
    #         sampleNumbers.append(int(fileNamesSplitted[2]))
    #         sampleMeasures.append(fileNamesSplitted[3])
    #     else:
    #         sampleNumbers.append(int(fileNamesSplitted[3]))
    #         sampleMeasures.append(fileNamesSplitted[4])
    # sampleNames = list(set(sampleNames))
    # times = list(set(times))
    # sampleNumbers = list(set(sampleNumbers))
    # sampleMeasures = list(set(sampleMeasures))
    # print(sampleMeasures)
    # sampleMeasures.sort(key=int)
    # print(sampleMeasures)
    # times.sort(key=int)
    # print(times)
    # for n in sampleNames:
    #     for i in sampleNumbers:
    #         for m in sampleMeasures:
    #             sheetNames.append('PSC_'+n+'_'+str(i)+'_'+m)
    # print(sheetNames)
    # WorkbookLayout.WorkbookLayout.makeSheetsPsc(wb, graphNames, sheetNames)
    #
    # for sn in sheetNames:
    #     activeSheet = wb[sn]
    #     for t in times:
    #         #pad+n-tmin-i-m.dat
    #         n = sn.split('_')[1]
    #         i = sn.split('_')[2]
    #         m = sn.split('_')[3]
    #         file_path = filePath + n + '-' + str(t) + 'min-' + str(i) + '-' + str(m) + '.dat'
    #         if os.path.exists(file_path):
    #             iv = Data.Data.getDataListPsc(file_path)
    #             WorkbookLayout.WorkbookLayout.setIVPsc(activeSheet, t, iv)
    #     print(sn)
    #     print(activeSheet.max_row)
    #     print(activeSheet.max_column)
    #     if(activeSheet.max_column==1 and activeSheet.max_row==1):
    #         wb.remove(activeSheet)
    #         print('sheet removed')
    #         continue
    #     Grafieken.Grafieken.makeChartPsc(sn, wb, times)
    # print('Saving...')
    # wb.save(filePath + 'test' + wbName)


    #----------------------------------- BIFI -----------------------------------

    # def begin(self, pv, maxHour, amount, sides, wbName, path):
    #     return True
    maxHour=1000
    wbName='PID222.xlsx'#'__PII_BIFI_npert_JW_5BB.xlsx'
    path='C:\\Users\Melissa\Documents\\01 Masterproef imec\Engineering\Python script\PID222-20190226T135803Z-001\PID222\\'#'C:\\Users\Melissa\Documents\\01 Masterproef imec\Engineering\Python script\SOLMAT-20190129T121844Z-001\SOLMAT\\'
    # sheetNames = ['SM_BSG1_F','SM_BSG1_R','SM_BSG2_F','SM_BSG2_R','SM_GBS1_F','SM_GBS1_R','SM_GBS2_F','SM_GBS2_R','SM_GG1','SM_GG1_R','SM_GG2_F','SM_GG2_R']

        # sheetNames = []
        # for n in range(1, int(amount) + 1, 1):
        #    if sides == []:
        #        sheetNames.append(pv + str(n) + '_Best') # -------------- -_Best
        #    else:
        #        for i in range(0, len(sides), 1):
        #            sheetNames.append(pv + str(n) +'_' + str(sides[i]))
    graphNames = ['%PID', 'FF', 'Voc', 'Isc']

    subfolders = [f.name for f in os.scandir(path) if f.is_dir()]
    subfolders.sort(key=float)
    sheetNames = [f.name for f in os.scandir(path + subfolders[len(subfolders) - 1]) if f.is_dir()]
    print(sheetNames)
    print(len(sheetNames))
    wb = openpyxl.load_workbook(path + wbName, data_only=True)
    if os.path.exists(path + subfolders[len(subfolders) - 1]):
        for f in listdir(path + subfolders[len(subfolders) - 1]):
            if f.startswith('data-exchange') and f.endswith('.xlsx'):
                data_exchange = f
                dataEx = openpyxl.load_workbook(path + subfolders[len(subfolders) - 1] + '/' + data_exchange)
                WorkbookLayout.WorkbookLayout.makeSheets(wb, dataEx, graphNames, sheetNames)
                info = AlgemeneInfo.AlgemeneInfo.datasheet(wb, sheetNames, dataEx, maxHour, subfolders)
                begin = info[0]
                hours = info[1]
                print('hours : ' + str(hours))
    else:
        print('Error')
        #return False
    # hours = [0, 2, 4, 6, 20, 42, 64, 91, 156, 309, 517, 539, 626]
    # print(hours)
    # print(subfolders)


    wbNameImg = wbName.replace('.xlsx','')
    #wb.save(path + wbName)
    Photo.Photo.makeTitle(wbNameImg)

    for n in range(0, len(sheetNames), 1):
        Photo.Photo.makeTitle(sheetNames[n])
        #begin = eerste rij die ingevult moet worden (dus begin-1), maar de hours beginnen pas op rij 2 (dus nog eens -1) => begin-2
        ivHours = []
        eqeHours = []
        for hour in range(begin-2, len(hours), 1):
            nieuwPath = path + str(subfolders[hour]) + '/' + sheetNames[n]
            photoPath = path + str(subfolders[hour]) + '/'
            """
            for f in listdir(photoPath):
                if (f.endswith(str(sheetNames[n]) + '.JPG') or f.endswith(str(sheetNames[n]) + '.jpg')):
                    photo = f
                    Photo.Photo.resize(photoPath + photo, photoPath + sheetNames[n] + '_resised')

            #HORIZONTALE TOEVOEGING
            if n == 0:
                Photo.Photo.makeTitle(str(hours[hour]) + ' h')
                Photo.Photo.merge_image(str(wbNameImg) + '.jpg', str(hours[hour]) + ' h.jpg', False, str(wbNameImg))
                os.remove(str(hours[hour]) + ' h.jpg')

            Photo.Photo.merge_image(str(sheetNames[n]) + '.jpg', photoPath + sheetNames[n] + '_resised.jpg', False, str(sheetNames[n]))

            #remove files
            os.remove(photoPath + sheetNames[n] + '_resised.jpg')
            """
            activeSheet = wb[sheetNames[n]]

            ivPath = nieuwPath + '/IV/'
            eqePath = nieuwPath + '/IQE/'

            drk = []
            lgt = []

            #data uit file halen
            if os.path.exists(ivPath):
                if os.path.exists(ivPath + sheetNames[n] + '.drk'):
                    drk = Data.Data.getDataList(str(ivPath) + sheetNames[n] + '.drk')
                else:
                    print(str(ivPath) + '.drk file bestaat niet')
                if os.path.exists(ivPath + sheetNames[n] + '.lgt'):
                    lgt = Data.Data.getDataList(str(ivPath) + sheetNames[n] + '.lgt')
                else:
                    print(str(ivPath) + sheetNames[n] + ' .lgt file bestaat niet')
            else:
                print(str(ivPath) + ' bestaat niet')

            if drk != [] and lgt != []:
                WorkbookLayout.WorkbookLayout.setIV(activeSheet, hours[hour], drk, lgt)
                ivHours.append(hours[hour])

            # ------ EQE ------
            eqeSheet = 'EQE_' + sheetNames[n]
            activeSheet = wb[eqeSheet]

            #data uit file halen
            eqeFile = ''
            if os.path.exists(eqePath):
                for f in listdir(eqePath):
                    if f.startswith(sheetNames[n]) and f.endswith('.eqe'):
                        eqeFile = f
                if eqeFile != '':
                    eqe = Data.Data.getDataList(str(eqePath) + eqeFile)[0]
                    WorkbookLayout.WorkbookLayout.setEQE(activeSheet, hours[hour], eqe)
                    eqeHours.append(hours[hour])
                else:
                    print(str(eqePath) + ' file bestaat niet')
            else:
                print(str(eqePath) + ' bestaat niet')
        """
        #VERTICALE TOEVOEGING
        Photo.Photo.merge_image(str(wbNameImg) + '.jpg', str(sheetNames[n]) + '.jpg', True, str(wbNameImg))
        os.remove(str(sheetNames[n]) + '.jpg')
        """
        #grafieken
        print(ivHours)
        print(eqeHours)
        Grafieken.Grafieken.makeChart(sheetNames[n], wb, ivHours)
        Grafieken.Grafieken.makeChart('EQE_' + sheetNames[n], wb, eqeHours)
    Grafieken.Grafieken.makeSeperateGraphs(wb, graphNames, sheetNames, hours)

    print('Saving...')
    wb.save(path + str(hours[len(hours) - 1]) + wbName)
    #return True

