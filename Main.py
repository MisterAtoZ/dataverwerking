import AlgemeneInfo
import IV
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

class Main():
    uren = AlgemeneInfo.AlgemeneInfo.datasheet()
    wb = openpyxl.load_workbook('__PID_BIFI_NPERT_JW_5BB_updated.xlsx', data_only=True)
    generalSheet = wb['General']

    for n in range(1, 2, 1):
        if (n == 0):
            sheet = 'JW1_B'
        elif (n == 1):
            sheet = 'JW1_F'
        elif (n == 2):
            sheet = 'JW2_B'
        else:
            sheet = 'JW2_F'

        drk = IV.IV.getIVlist(str(sheet) + '.drk', sheet)
        lgt = IV.IV.getIVlist(str(sheet) + '.lgt', sheet)

        vDark = drk[0]
        iDark = drk[1]
        vLight = lgt[0]
        iLight = lgt[1]

        activeSheet = wb[sheet]

        #uur
        column1 = activeSheet.max_column + 2
        column2 = activeSheet.max_column + 5
        activeSheet.merge_cells(start_row = 1, start_column = column1, end_row = 1, end_column = column2)
        activeSheet.cell(row=1, column=column1).value = str(uren[0]) + ' h'


        #light dark
        column3 = column1 + 1
        activeSheet.merge_cells(start_row = 2, start_column = column1, end_row = 2, end_column = column3)
        activeSheet.cell(row=2, column=column1).value = 'light'

        column4 = column1 + 2
        column5 = column4 + 1
        activeSheet.merge_cells(start_row = 2, start_column = column4, end_row = 2, end_column = column5)
        activeSheet.cell(row=2, column=column4).value = 'dark'


        #iv invullen
        activeSheet.cell(row=3, column=column1).value = 'V'
        activeSheet.cell(row=3, column=column3).value = 'I'
        activeSheet.cell(row=3, column=column4).value = 'V'
        activeSheet.cell(row=3, column=column5).value = 'I'

        #data invullen




    print('Saving...')
    wb.save('__PID_BIFI_NPERT_JW_5BB_updated.xlsx')

