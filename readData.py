import  _thread
import serial
import datetime
from datetime import datetime
import time
import tkinter
import math   # This will import math module
import xlsxwriter
import signal
import sys
import pandas as pd

ser = serial.Serial("COM6", 19200)
f = open("demofile.txt", "w")
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()

values = []
start = int(round(time.time() * 1000))
counter = 0

def signal_handler(sig, frame):
        print('You pressed Ctrl+C!')
        # Create a Pandas dataframe from the data.
        df = pd.DataFrame(values)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        excel_file = 'Graph.xlsx'
        sheet_name = 'Sheet1'

        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'line'})

        # Configure the series of the chart from the dataframe data.

        chart.add_series({
        'categories': '=Sheet1!$F$2:$F$' + str(counter),
        'values':     '=Sheet1!$E$2:$E$' + str(counter),
        })

        # Configure the chart axes.
        chart.set_x_axis({'name': 'Time (ms)', 'position_axis': 'on_tick'})
        chart.set_y_axis({'name': 'cm/s^2', 'major_gridlines': {'visible': False}})

        # Turn off chart legend. It is on by default in Excel.
        chart.set_legend({'position': 'none'})

        # Insert the chart into the worksheet.
        worksheet.insert_chart('J2', chart)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        print("Done")
        sys.exit(0)
        

signal.signal(signal.SIGINT, signal_handler)
timeINseconds = 0
while True:
    # ts = datetime.now().microsecond
    ts = int(round(time.time() * 1000))
    timeINseconds = (ts - start)/1000
    cc=str(ser.readline())
    toSplit = cc[2:][:-5]
    array = toSplit.split(".")
    values.append([array[0] , array[1] , array[2] , int(math.sqrt(int(array[0])**2 +  int(array[1])**2 +  int(array[2])**2)) , timeINseconds])
    counter = counter + 1
    print(array[0] + " " + array[1] + " " + array[2] + " " +  str(int(math.sqrt(int(array[0])**2 +  int(array[1])**2 +  int(array[2])**2)))  + " "  + str(timeINseconds))
#     worksheet.write(row, 0, array[0])
#     worksheet.write(row, 1, array[1])
#     worksheet.write(row, 2, array[2])
#     worksheet.write(row, 3, int(math.sqrt(int(array[0])**2 +  int(array[1])**2 +  int(array[2])**2)))
#     worksheet.write(row, 4, int(ts) - int(start))
#     row += 1
