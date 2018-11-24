import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment

class WorkbookLayout():
    def makeSheets(workbook, dataEx, graphNames, sheetNames):
        wb = workbook
        data = dataEx['data-exchange']
        # fonts
        vet = Font(bold=True)
        groot = Font(size=14)

        for j in range(0,len(graphNames),1):
            if graphNames[j] not in wb.sheetnames:
                wb.create_sheet(graphNames[j])

        for n in range(0, len(sheetNames), 1):
            if sheetNames[n] not in wb.sheetnames:
                wb.create_sheet(sheetNames[n])
                sheet = wb[sheetNames[n]]
                sheet.cell(row=1, column=1).value = 'Time [h]'
                sheet.cell(row=1, column=1).font = vet
                for j in range(2, 13, 1):
                    sheet.cell(row=1, column=j).value = data.cell(row=2, column=j + 7).value + ' [' + data.cell(row=3, column=j + 7).value + ']'
                    sheet.cell(row=1, column=j).font = vet
                sheet.cell(row=1, column=17).value = '%PID [%]'
                sheet.cell(row=1, column=17).font = vet
            if 'EQE_' + sheetNames[n] not in wb.sheetnames:
                wb.create_sheet('EQE_' + sheetNames[n])
                sheet = wb['EQE_' + sheetNames[n]]
                sheet.cell(row=2, column=1).value = 'wavelength (Lambda)'
                for j in range(0,93,1):
                    sheet.cell(row=j+3, column=1).value = 280 + j*10


    def setIV(activeSheet, uur, drk, lgt):
        # fonts
        vet = Font(bold=True)
        groot = Font(size=14)

        vDark = drk[0]
        iDark = drk[1]
        vLight = lgt[0]
        iLight = lgt[1]

        # uur
        column1 = activeSheet.max_column + 2
        column2 = activeSheet.max_column + 5
        activeSheet.merge_cells(start_row=1, start_column=column1, end_row=1, end_column=column2)
        activeSheet.cell(row=1, column=column1).value = str(uur) + ' h'
        activeSheet.cell(row=1, column=column1).font = vet
        activeSheet.cell(row=1, column=column1).alignment = Alignment(horizontal='center')

        # light dark
        column3 = column1 + 1
        activeSheet.merge_cells(start_row=2, start_column=column1, end_row=2, end_column=column3)
        activeSheet.cell(row=2, column=column1).value = 'light'
        activeSheet.cell(row=2, column=column1).font = groot
        activeSheet.cell(row=2, column=column1).alignment = Alignment(horizontal='center')

        column4 = column1 + 2
        column5 = column4 + 1
        activeSheet.merge_cells(start_row=2, start_column=column4, end_row=2, end_column=column5)
        activeSheet.cell(row=2, column=column4).value = 'dark'
        activeSheet.cell(row=2, column=column4).font = groot
        activeSheet.cell(row=2, column=column4).alignment = Alignment(horizontal='center')

        # iv invullen
        activeSheet.cell(row=3, column=column1).value = 'V'
        activeSheet.cell(row=3, column=column3).value = 'I'
        activeSheet.cell(row=3, column=column4).value = 'V'
        activeSheet.cell(row=3, column=column5).value = 'I'

        # data invullen
        # --------------- verschil lengte light / dark ? ---------------
        for j in range(0, len(vLight), 1):
            activeSheet.cell(row=j + 4, column=column1).value = vLight[j]
            activeSheet.cell(row=j + 4, column=column3).value = iLight[j]
        for j in range(0, len(vDark), 1):
            activeSheet.cell(row=j + 4, column=column4).value = vDark[j]
            activeSheet.cell(row=j + 4, column=column5).value = iDark[j]


    def setEQE(activeSheet, uur, eqe):
        # titels invullen
        column1 = activeSheet.max_column + 2
        column2 = activeSheet.max_column + 3
        activeSheet.merge_cells(start_row=1, start_column=column1, end_row=1, end_column=column2)
        activeSheet.cell(row=1, column=column1).value = str(uur) + ' h'

        activeSheet.cell(row=2, column=column1).value = 'EQE'
        activeSheet.cell(row=2, column=column2).value = 'EQE norm.'

        # data invullen
        for j in range(0, len(eqe), 1):
            activeSheet.cell(row=j + 3, column=column1).value = eqe[j]
            activeSheet.cell(row=j + 3, column=column2).value = eqe[j] / activeSheet.cell(row=j + 3, column=3).value
