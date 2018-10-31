import openpyxl
import sys
from openpyxl.utils import get_column_letter, column_index_from_string

class IV():
    wb = openpyxl.load_workbook('__PID_BIFI_NPERT_JW_5BB.xlsx', data_only=True)
    sheet = wb['JW1_F']

    def getIVlist(filename):
        with open(filename, 'r') as file:
            wholeFile = file.read()
            data = wholeFile.split('**Data**')[1]
            # print(data)

            splitted = data.split()

            v = []
            i = []

            if str(filename).endswith('drk'):
                jump = 3
            else :
                jump = 4

            for j in range(0, len(splitted),jump):
                v.append(splitted[j])
                i.append(splitted[j+1])

            #print(v)
            #print(i)

            iv = [v, i]

            return iv