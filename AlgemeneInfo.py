import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

class AlgemeneInfo():
    # action = input("wat is de file waar de data in moet komen te staan? naam.xlsx")
    wb = openpyxl.load_workbook('__PID_BIFI_npert_JW_5BB.xlsx', data_only=True)
    uren = input('hoeveel uren is er gestressed?: ')
    print(uren)

    dataNaam = "data-exchange"+uren+".xlsx"
    print(dataNaam)
    dataEx = openpyxl.load_workbook(dataNaam)

    sheet = wb['JW1_B']
    data = dataEx['data-exchange']
    generalSheet = wb['General']

    aantalMetingen = 0
    for i in range(19, generalSheet.max_row, 1):
        if generalSheet.cell(row=i, column=5).value == None :
            break
        else :
            aantalMetingen = aantalMetingen + 1
    # print('aantalMetingen ' + str(aantalMetingen))

    nextRow = 0
    for cellObj in sheet['A']:
        if cellObj.value == None:
            nextRow = cellObj.row
            break
    #print('nextRow ' + str(nextRow))
    begin = nextRow

    for rows in range(begin, 2+aantalMetingen, 1):
        sheet.cell(row=nextRow, column=1).value = '=General!E' + str(17+nextRow)

        #print(sheet.cell(row=nextRow, column=1).value + " " +str(nextRow))
        #print(sheet)        #print(sheet==wb['JW1_F'])

        if sheet == wb['JW2_B']:
            staal_nr = 4
        elif sheet == wb['JW2_F']:
            staal_nr = 5
        if sheet == wb['JW1_B']:
            staal_nr = 6
        elif sheet == wb['JW1_F']:
            staal_nr = 7
        #print(staal_nr)

        uur = int(round(float(generalSheet.cell(row=17+nextRow, column=5).value),0))
        print(uur)
        for n in range(0,3,1):
            if(n==0):
                sheet = wb['JW1_B']
            elif (n==1):
                sheet = wb['JW1_F']
            elif (n==2):
                sheet = wb['JW2_B']
            else:
                sheet = wb['JW2_F']

            for i in range(staal_nr, data.max_row, 4):
                print(str(data.cell(row=i, column=1).value) + " " + str(uur))
                if(str(data.cell(row=i, column=1).value) == str(uur)):
                    for j in range(2, 13, 1):
                        sheet.cell(row=nextRow, column=j).value = data.cell(row=i, column=j+7).value
                        print(sheet.cell(row=nextRow, column=j).value)
                    break

            nextRow = nextRow + 1
        print('-----------------------')
        print('Saving...')
        wb.save('__PID_BIFI_NPERT_JW_5BB_updated.xlsx')
