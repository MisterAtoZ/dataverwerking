import openpyxl

class Grafieken():
    def makeChart(sheetname, workbook, uren):
        wb = workbook
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

    def makeSeperateGraphs(workbook, sheetNames, uren):
        wb = workbook
        pid = wb['%PID']
        ff = wb['FF']
        voc = wb['Voc']
        isc = wb['Isc']

        graphs = [pid, ff, voc, isc]

        for j in range(0,len(graphs),1):
            chartObj = openpyxl.chart.ScatterChart()
            chartObj.legend.position = 'b'

            if j == 0:
                chartObj.x_axis.title = 'Time [h]'
                chartObj.y_axis.title = '%PID [%]'
                chartObj.y_axis.scaling.max = 100
                x = 1
                y = 17
            else:
                chartObj.x_axis.title = '%PIDs [%]'
                chartObj.x_axis.scaling.max = 100
                x = 17
                if j == 1:
                    chartObj.y_axis.title = 'Fill Factor [%]'
                    y = 8
                if j == 2:
                    chartObj.y_axis.title = 'Voc [mV]'
                    y = 7
                if j == 3:
                    chartObj.y_axis.title = 'Current [mA]'
                    y = 5

            for i in range(0, len(sheetNames), 1):
                xvalues = openpyxl.chart.Reference(wb[sheetNames[i]], min_col=x, min_row=2, max_col=x, max_row=len(uren)+1)
                yvalues = openpyxl.chart.Reference(wb[sheetNames[i]], min_col=y, min_row=2, max_col=y, max_row=len(uren)+1)
                seriesObj = openpyxl.chart.Series(yvalues, xvalues, title=str(sheetNames[i]))
                seriesObj.marker = openpyxl.chart.marker.Marker('circle')
                chartObj.append(seriesObj)

            graphs[j].add_chart(chartObj)
