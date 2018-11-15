import os
class IV():
     def getIVlist(filename, sheetname):
        with open(filename, 'r') as file:
            wholeFile = file.read()
            data = wholeFile.split('**Data**')[1]
            #print(data)

            splitted = data.split()

            v = []
            i = []

            if str(filename).endswith('drk'):
                jump = 3
            elif str(filename).endswith('lgt'):
                jump = 4
            else:
                jump = 2

            for j in range(0, len(splitted),jump):
                v.append(float(splitted[j]))
                i.append(float(splitted[j+1]))


            #print(v)
            #print(i)

            iv = [v, i]

            return iv