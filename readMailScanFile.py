# import the required libraries
from __future__ import print_function
import pickle
import os.path
import io
import shutil
import pandas as pd
import os
from os import listdir
from os.path import isfile, join
from datetime import datetime
from datetime import timedelta

class DriveAPI:
    global SCOPES

    def __init__(self) -> None:
        txtPath = os.getcwd() + "/File/*"
        os.system("rm -rf "+txtPath)


    def LocateTotal(self):
        namesInFile = [
            "jgw",
            "newstart",
            "prudent",
            "nubreak",
            "pointbreak",
            "clearone",
            "secureone",
            "safestone"
        ]

        pathName = os.getcwd() + "/mailscan/"

        onlyfiles = [f for f in listdir(pathName) if isfile(join(pathName, f))]

        fileNames = []
        for file in onlyfiles:
            for name in namesInFile:
                if name in file.lower():
                    fileNames.append(file)

        for fileName in fileNames:
            print(fileName)

            # Get current data to subtract and find the correct sheet
            now = datetime.now()
            today_weekday = now.weekday()
            # print("weekday: "+ str(today_weekday))
            # 0 -> Monday

            count_days = 14
            # find wednesday of the current week
            if today_weekday == 0:
                count_days -= 2
            elif today_weekday == 1:
                count_days -= 1
            elif today_weekday == 2:
                count_days += 0
            elif today_weekday == 3:
                count_days += 1
            elif today_weekday == 4:
                count_days += 2
            elif today_weekday == 5:
                count_days += 3
            elif today_weekday == 6:
                count_days += 4
            else:
                print("invalid weekday")
                exit(0)

            correctSheet = now - timedelta(days=count_days)
            previousSheet = correctSheet - timedelta(days=7)
            correctSheet = correctSheet.strftime("%m%d%Y")
            previousSheet = previousSheet.strftime("%m%d%Y")
            print("Sheet: " + correctSheet)
            print("Previous Sheet: " + previousSheet)

            xlsx_file = pd.read_excel("mailscan/"+fileName, sheet_name = None)
            dates = []

            for date in xlsx_file.keys():
                # print("sheet: " + str(date))
                dates.append(date)

            sheetDates = []
            sheetDates.append(previousSheet)
            sheetDates.append(correctSheet)
            
            msg = ''
            datesAdded = 0
            for sheetDate in sheetDates:
                if len(dates) == 0 or sheetDate not in dates:
                    print("Valid Sheet not detected")
                    if "safestone" in fileName.lower() and not datesAdded:
                        datesAdded = 1

                        correctSheet = now - timedelta(days=count_days-1)
                        previousSheet = correctSheet - timedelta(days=7)
                        correctSheet = correctSheet.strftime("%m%d%Y")
                        previousSheet = previousSheet.strftime("%m%d%Y")

                        # sheetDates = []
                        sheetDates.append(previousSheet)
                        sheetDates.append(correctSheet)

                        print("Sheet added: " + correctSheet)
                        print("Sheet2 added: " + previousSheet)
                    continue

                # find Total field and it's percentage
                found = 0
                for i, j in xlsx_file[sheetDate].iterrows():
                    for idx,row in enumerate(j):
                        if row == "Total" or found == 1:
                            found = 1
                            if pd.isna(row):
                                break
                            percent = row

                    if found:
                        break

                percent = round(percent*100,2)
                msg += sheetDate[0:2] + "/" + sheetDate[2:4] + " drop is " + str(percent) + "% scanned.\n"
            print(msg)

            saveTxtFile(fileName, msg)


def saveTxtFile(fileName, msg):
    onlyFileName = fileName.split("/")[-1].replace("xlsx","txt")

    with open("File/"+onlyFileName, 'w') as f:
        f.write(msg)
    print("File txt ready\n\n")


if __name__ == "__main__":
    obj = DriveAPI()
    obj.LocateTotal()
