class Data():

    def getDataList(filename):
        """
        read the file and get the I and V values and return it
        :param filename: complete path to the file
        :return: list containing a list of the I values and a list of the V values
        """
        with open(filename, 'r') as file:
            wholeFile = file.read()
            data = wholeFile.split('**Data**')[1]
            print(filename)
            splitted = data.split()
            v = []
            i = []


            # the IV parameters are in columns after '**Data**' in the file
            # in the .drk files, 3 columns are present in the file, only the first 2 columns are important
            # in the .lgt files, 4 columns are present in the file, only the first 2 columns are important
            # in other files, only 2 columns are present
            if str(filename).endswith('drk'):
                jump = 3
            elif str(filename).endswith('lgt'):
                jump = 4
            else:
                jump = 2

            for j in range(0, len(splitted),jump):
                if(j+1 < len(splitted)):
                    v.append(float(splitted[j]))
                    i.append(float(splitted[j+1]))

            iv = [i, v]
            return iv

    def getDataListPsc(filename):
        with open(filename, 'r') as file:
            data = file.read()
            splitted = data.split()
            v = []
            i = []

            jump = 2

            for j in range(0, len(splitted), jump):
                if(j+1<len(splitted)):
                    # print('j: ' + str(j))
                    # print(splitted[j])
                    # print(splitted[j+1])
                    i.append(float(splitted[j]))
                    v.append(float(splitted[j+1]))

            iv = [i, v]

            return iv