'''
Метод заполняет requirements.txt
Заполняет exceptions.csv
'''


import os
import csv
import re

from pylint import epylint


folders = [os.getcwd() + d for d in ["\\Parsers20162017"]]
path = os.getcwd()


def check_files(directory, ipynb_to_py=False):
    """Checking all files and create report array:
    [{file: filename, code: errorcode, error: errorbody},]

    if ipynb_to_py == true method create .py file for each .ipynb file 
    """
    filelist = []
    exceptionlist = []
    
    for root, _, files in os.walk(directory):
        res = ""
        if "venv" in root or "__pycache__" in root or "check" in root:
            continue
        for file in files:
            filelist.append(f'{root}\\{file}')
        for file in filelist:
            if ".txt" in file or ".csv" in file:
                continue
            if ".ipynb" in file:
                if ipynb_to_py:
                    os.popen(f'jupyter nbconvert "{file}" --to python')
                    res = os.popen(f'nbqa pylint "{file}"').read()
                else:
                    continue
            elif ".py" in file:
                print(f"current file: {file}")
                res = os.popen(f'pylint "{file}"').read()
            if res == "":
                continue

            for row in res.split("\n"):
                listrow = row.split(":")
                if len(listrow) > 4:
                    code = listrow[3]
                    e = listrow[4]
                    if "E" in code:
                        exceptionlist.append({
                            "file": file,
                            "code": code,
                            "error": e,
                        })

    return exceptionlist


def save_to_csv(dir_name, filename, array):
    """Write data to csv"""
    with open(
        f'{dir_name}\\{filename}',
        "w", encoding="UTF-8",
        newline='\r\n'
    ) as csvfile:

        header = array[0].keys()
        writer = csv.DictWriter(f=csvfile, fieldnames=header, lineterminator="\n")
        writer.writeheader()

        for k in array:
            writer.writerow(k)


try:
    for directory in folders:
        save_to_csv(directory, "excpetions.csv", check_files(directory, True))
except Exception as e:
    print(e)
