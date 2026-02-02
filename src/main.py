import os
import pandas as pd
import openpyxl as op

from datetime import datetime

def main():

    for e in os.scandir(getTimesheetDirPath()):
        if e.is_file():
            sheet = loadSheet(e.path)

            timeCells = []
            new = []
            for r in range(6, 13):
                new = []
                for c in range(2, 10):
                    if sheet.cell(r, c).value != None:
                        new.append(sheet.cell(r, c).value)
                    else:
                        new.append(0)
                timeCells.append(new)
            
            printCells(timeCells)
            print()
            countWorkedHours(timeCells)
            print(sumHours(hours))
            print()
    print()

def loadSheet(path):
    script_dir = os.path.dirname(__file__) 
    rel_path = path
    abs_file_path = os.path.join(script_dir, rel_path)
    wb = op.load_workbook(abs_file_path)
    return wb.active
    
def printCells(cells):
    for y in range(len(cells)):
        match y:
            case 0:
                print("Sun:", end = ' ')
            case 1:
                print("Mon:", end = ' ')
            case 2:
                print("Tue:", end = ' ')
            case 3:
                print ("Wed:", end = ' ')
            case 4:
                print ("Thu:", end = ' ')
            case 5:
                print ("Fri:", end = ' ')
            case 6:
                print ("Sat:", end = ' ')
        for x in range(len(cells[0])):
            print(cells[y][x], end = ' ')
        print()

def countWorkedHours(cells):
    shifts = {}
    for day in range(7):
        print(f"{indToDay[day]}: ", end = ' ')
        for shift in range(0, 6, 2):
            if (cells[day][shift] != 0) and (cells[day][shift + 1] != 0):
                print(f"{cells[day][shift + 1].hour - cells[day][shift].hour}", end = ' ')
                for hr in range(cells[day][shift].hour, cells[day][shift + 1].hour):
                    hours[hr] += 1
        print()

def getTimesheetDirPath():
    script_dir = os.path.dirname(__file__) 
    rel_path = "sumtimesheets/Excel/"
    return os.path.join(script_dir, rel_path)

def sumHours(hours):
    return sum([hours[hr] for hr in hours])

def genIndToDayDict():
    d = {
        0 : "Sun",
        1 : "Mon",
        2 : "Tue",
        3 : "Wed",
        4 : "Thu",
        5 : "Fri",
        6 : "Sat"
    }
    return d

def genHourTimeDict():
    h = {}
    for i in range(24):
        if (i < 10):
            key = "{}:00:00".format("0" + str(i))
        else:
            key = "{}:00:00".format(i)
        h[key] = 0
    return h

def genHourDict():
    h = {}
    for i in range(24):
        h[i] = 0
    return h

indToDay = genIndToDayDict()
hours = genHourDict()

main()