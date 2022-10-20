'''
Метод заполняет requirements.txt
Заполняет exceptions.csv
'''

import os
import csv
import re

from pylint import epylint


folders = [os.getcwd() + d for d in ["\\Parsers20152017", "\\ParsersMesayhaNNG", "\\Parsersother"]]
path = os.getcwd()


def check_files(dir):
    print(dir)
    filelist = []
    exceptionlist = []
    modules_list = []
    
    for root, _, files in os.walk(dir):
        res = ""
        if "venv" in root or "__pycache__" in root or "check" in root:
            continue
        for file in files:
            print(f'{root}\\{file}')
            filelist.append(f'{root}\\{file}')
            # print(file)

        for file in filelist:
            if ".txt" in file:
                continue
            if ".ipynb" in file:
                res = os.popen(f'nbqa pylint "{file}"').read()
            elif ".py" in file:
                res = os.popen(f'pylint "{file}"').read()
            if res == "":
                continue

            for r in res.split("\n"):
                lr = r.split(":")
                if len(lr) > 5:
                    code = lr[4]
                    e = lr[5]
                    if "E" in code:
                        exceptionlist.append({
                            "file": file,
                            "code": code,
                            "error": e,
                        })
                        # print(exceptionlist)

    return exceptionlist


def save_to_csv(dir, filename, array):
    with open(f'{dir}\\{filename}', "w", encoding="UTF-8", newline='\r\n') as csvfile:
        header = array[0].keys()
        writer = csv.DictWriter(f=csvfile, fieldnames=header, lineterminator="\n")
        writer.writeheader()
        for k in array:
            writer.writerow(k)

try:
    for dir in folders:
        save_to_csv(dir, "excpetions.csv", check_files(dir))
except Exception as e:
    print(e)