import csv
import sys
from excel_output_utils import write_collected_data_to_worksheet, save_workbook
import os

for filename in os.listdir(os.getcwd()):
    if '.csv' in filename:
        print('converting',filename)
        try:
            csv_name = filename

            csv_data = []
            csv_file = open(csv_name, newline='\n')
            opened_csv = csv.reader(csv_file, delimiter=',', quotechar='\'')
            for row in opened_csv:
                csv_data.append([row])

            write_collected_data_to_worksheet(csv_name.replace('.csv',''),None,csv_data)
            save_workbook(csv_name.replace('.csv','')+'_converted')
        except:
            print('Ouch! Something is wrong')
