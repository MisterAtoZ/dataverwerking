import openpyxl

class Grafieken():
     def makeChart(sheetname, workBook, uren):
        wb = workBook
        print(wb)
        print(sheetname)
        sheet = wb[sheetname]


        if sheetname.startswith('EQE'):
            for j in range(0,2,1):
                chartObj = openpyxl.chart.ScatterChart()
                chartObj.x_axis.title = 'Wavelenght [nm]'
                chartObj.legend.position = 'b'
                if j == 0:
                    chartObj.y_axis.title = 'EQE [-]'
                    location = 'C20'
                else:
                    chartObj.y_axis.title = 'Normalized EQE [-]'
                    location = 'M20'

                for i in range(0,len(uren),1):
                    xvalues = openpyxl.chart.Reference(sheet, min_col=1, min_row=3, max_col=1, max_row=sheet.max_row)
                    yvalues = openpyxl.chart.Reference(sheet, min_col=3+(i*3)+j, min_row=3, max_col=3+(i*3)+j, max_row=sheet.max_row)
                    seriesObj = openpyxl.chart.Series(yvalues, xvalues, title=str(uren[i]) + ' h')
                    chartObj.append(seriesObj)

                sheet.add_chart(chartObj, location)

        else:
            for j in range(0,2,1):
                chartObj = openpyxl.chart.ScatterChart()
                chartObj.x_axis.title = 'Voltage [V]'
                chartObj.y_axis.title = 'Current [A]'
                chartObj.x_axis.scaling.min = 0
                chartObj.legend.position = 'b'
                if j == 0:
                    chartObj.title = 'Light IV'
                    chartObj.y_axis.scaling.max = 0;
                    location = 'C20'
                else:
                    chartObj.title = 'Dark IV'
                    location = 'J20'

                for i in range(0,len(uren),1):
                    xvalues = openpyxl.chart.Reference(sheet,min_col=20+(i*5)+(j*2),min_row=4,max_col=20+(i*5)+(j*2),max_row=sheet.max_row)
                    yvalues = openpyxl.chart.Reference(sheet,min_col=21+(i*5)+(j*2),min_row=4,max_col=21+(i*5)+(j*2),max_row=sheet.max_row)
                    seriesObj = openpyxl.chart.Series(yvalues, xvalues, title=str(uren[i]) + ' h')
                    chartObj.append(seriesObj)

                sheet.add_chart(chartObj, location)





"""

        chartObj = openpyxl.chart.ScatterChart()
        chartObj.title = 'Light IV'
        chartObj.x_axis.title = 'Voltage [V]'
        chartObj.y_axis.title = 'Current [A]'

        #i(v) = colom2(colom1)
        for j in range(0,k,1):
            for i in range(0,len(uren),1):
                xvalues = openpyxl.chart.Reference(sheet, min_col=20+(i*5), min_row=4, max_col=20+(i*5), max_row=sheet.max_row)
                yvalues = openpyxl.chart.Reference(sheet, min_col=21+(i*5), min_row=4, max_col=21+(i*5), max_row=sheet.max_row)
                seriesObj = openpyxl.chart.Series(yvalues, xvalues, title=str(uren[i]) + ' h')
                chartObj.append(seriesObj)

            sheet.add_chart(chartObj, 'C20')

"""