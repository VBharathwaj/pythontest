import os
import logging
from configparser import SafeConfigParser
import openpyxl

def getColumnIndex(nameIndex):
    switcher = { 
        'A': 0, 
        'B': 1, 
        'C': 2,
        'D': 3, 
    } 
    return switcher.get(nameIndex, "nothing") 

logging.basicConfig(filename='status.log',level=logging.DEBUG, format='%(asctime)s %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p')
cur_wrk_dir=os.getcwd()

parser = SafeConfigParser()
parser.read(cur_wrk_dir + "\\input.ini")

logging.info("------------------------------------------------------------------------------------------")
logging.info("------------------------------------Started Processing------------------------------------")
logging.info("------------------------------------------------------------------------------------------")
logging.info("Fetching Input")

document_name = parser.get('input','document_name')
tab_name = parser.get('input','tab_name')
column_name = parser.get('input','column_name')
search_text = parser.get('input','search_text')

logging.info("Document Name"+ document_name)
logging.info("Tab Name"+ tab_name)
logging.info("Column Name"+ column_name)
logging.info("Search Text"+ search_text)

logging.info("Reading Input File")
wb = openpyxl.load_workbook(document_name)

sheet=wb[tab_name]
header= 0
column_data = []

column_name_index = column_name.split(" ")[1]
column_index = getColumnIndex(column_name_index)

for row in sheet:
    if header==0:
        header+=1
        continue
    column_data.append(str(row[column_index].value))

flag = False
for single_data in column_data:
    if single_data == search_text:
        flag = True
        break

if(flag):
    print("Found")
else:
    print("Not Found")