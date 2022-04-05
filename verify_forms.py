import os
import re
import shutil

import openpyxl

cwd = os.getcwd()
os.makedirs('{}/ОШИБКА'.format(cwd), exist_ok=True)
slovar = {
    'OPK1':'F18',
    'OPK2':'C16',
    'OPK3':'N12',
    'OPK4':'F12',
    'OPK5':'D19',
    'OPK6':'R12',
}

for file in os.listdir(cwd):
    for key, ya in slovar.items():
        if file.startswith(key):
            print(cwd + '/' + file)
            wb = openpyxl.reader.excel.load_workbook(filename=file, data_only=True)
            wb.active = 0
            sheet = wb.active
            a = sheet[str(ya)].value
            b = re.findall(r'\d{8}', file)
            try:
                int(a)
                int(b[0])
                int(a) == int(b[0])
            except:
                source_path = cwd + '/' + file
                destination_path = cwd + '/' + 'ОШИБКА' + '/' + file
                new_location = shutil.move(source_path, destination_path)
