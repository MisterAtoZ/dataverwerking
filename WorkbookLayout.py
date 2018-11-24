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
                print('sheet created')

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
                print('sheet created')
            if 'EQE_' + sheetNames[n] not in wb.sheetnames:
                wb.create_sheet('EQE_' + sheetNames[n])
                sheet = wb['EQE_' + sheetNames[n]]
                sheet.cell(row=2, column=1).value = 'wavelength (Lambda)'
                for j in range(0,93,1):
                    sheet.cell(row=j+3, column=1).value = 280 + j*10
                print('eqe sheet created')
