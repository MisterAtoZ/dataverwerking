import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

class AlgemeneInfo():
    def datasheet():
        wb = openpyxl.load_workbook('__PID_BIFI_NPERT_JW_5BB.xlsx', data_only=True)
        dataEx = openpyxl.load_workbook('data-exchange752.xlsx')

        data = dataEx['data-exchange']
        generalSheet = wb['General']

        # uren berekenen
        vorigUur = 0
        uren = [0]
        for i in range(20, generalSheet.max_row, 1):
            if generalSheet.cell(row=i, column=1).value == None:
                break
            else:
                date1 = generalSheet.cell(row=i, column=1).value
                print(date1)
                date2 = generalSheet.cell(row=i - 1, column=1).value
                timeOut = generalSheet.cell(row=i, column=2).value
                timeIn = generalSheet.cell(row=i - 1, column=3).value
                print(timeOut)

                date1 = date1.replace(hour=timeOut.hour, minute=timeOut.minute)
                print(date1)
                date2 = date2.replace(hour=timeIn.hour, minute=timeIn.minute)
                print(date2)

                seconds = (date1 - date2).total_seconds()
                interval = seconds / 3600

                generalSheet.cell(row=i, column=4).value = interval
                generalSheet.cell(row=i, column=5).value = interval + vorigUur

                vorigUur = generalSheet.cell(row=i, column=5).value
                uren.append(vorigUur)

        for n in range(0, 4, 1):
            if (n == 0):
                sheet = wb['JW1_B']
            elif (n == 1):
                sheet = wb['JW1_F']
            elif (n == 2):
                sheet = wb['JW2_B']
            else:
                sheet = wb['JW2_F']

            #print('----NEXT ROW----')
            nextRow = 0
            for cellObj in sheet['A']:
                if cellObj.value == None:
                    nextRow = cellObj.row
                    break
            #print('nextRow ' + str(nextRow))
            begin = nextRow


            #print('----INVULLEN----')
            for rows in range(begin, 2+len(uren), 1):
                print(str(len(uren)))
                uurInt = int(round(uren[rows-2],0))
                uurAfgerond = round(uren[rows-2],2)
                #print('uren[' + str(rows) + '] ' + str(uur))
                sheet.cell(row=nextRow, column=1).value = uurAfgerond

                #print(str(sheet.cell(row=nextRow, column=1).value) + " " +str(nextRow))

                if sheet == wb['JW2_B']:
                    staal_nr = 4
                elif sheet == wb['JW2_F']:
                    staal_nr = 5
                if sheet == wb['JW1_B']:
                    staal_nr = 6
                elif sheet == wb['JW1_F']:
                    staal_nr = 7
                #print(staal_nr)

                for i in range(staal_nr, data.max_row, 4):
                    #print(str(data.cell(row=i, column=1).value) + " " + str(uurInt))
                    if(str(data.cell(row=i, column=1).value) == str(uurInt)):
                        for j in range(2, 13, 1):
                            sheet.cell(row=nextRow, column=j).value = data.cell(row=i, column=j+7).value
                            #print(sheet.cell(row=nextRow, column=j).value)
                        break

                nextRow = nextRow + 1
                #print('-----------------------')

        print('Saving...')
        wb.save('__PID_BIFI_NPERT_JW_5BB_updated.xlsx')
        for i in range(0, len(uren),1):
            uren[i] = int(round(uren[i],0))

        return uren