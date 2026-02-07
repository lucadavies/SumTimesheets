import os
import openpyxl as op
import plotly as pt

timesheetsLocation = "sumtimesheets/Excel/"
debugCellRead = False
debugHourCount = False

def main():

    # Iterate over all files in directory specified
    for file in os.scandir(getTimesheetDirPath()):
        if file.is_file():
            if debugCellRead or debugHourCount:
                print(file.name)
            sheet = loadSheet(file.path)
            timeCells = getTimeCells(sheet)
            
            if debugCellRead:
                printCells(timeCells)
                print()
            countWorkedHours(timeCells)
    print(f"Total hours: {sumHours(hours)}")
    print()

""" Load specific sheet from workbook at provided path. """
def loadSheet(path):
    script_dir = os.path.dirname(__file__) 
    rel_path = path
    abs_file_path = os.path.join(script_dir, rel_path)
    wb = op.load_workbook(abs_file_path)
    return wb.active

""" Reads relevant cells from provided sheet and returns them in a 2D array. """
def getTimeCells(sheet):
    timeCells = [] # Directly holds cells containing times from timesheet

    # Rows 6 thru 12
    for r in range(6, 13):
        new = [] # Holding var for latest row of cells read in

        # Columns B thru I (regular shifts)
        for c in range(2, 10):
            if sheet.cell(r, c).value != None:
                new.append(sheet.cell(r, c).value)
            else:
                new.append(0) # Read empty cells as zeros

        # Column N (get-out)
        if sheet.cell(r, 14).value != None:
            new.append(sheet.cell(r, 14).value)
        else:
            new.append(0)
        timeCells.append(new)

    return timeCells

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

""" Takes 2D array containing cells read from timesheet and maps each hour worked to the hours dictionary 12am thru 11pm"""
def countWorkedHours(cells):
    for day in range(7):
        if debugHourCount:
            print(f"{indToDay[day]}: ", end = ' ')

        # For each shift start/end time pair...
        for shift in range(0, 8, 2):

            # Check shift has both a start AND end time
            if (cells[day][shift] != 0) and (cells[day][shift + 1] != 0):
                if debugHourCount:
                    print(f"{cells[day][shift + 1].hour - cells[day][shift].hour}", end = ' ')
                # For each hour spanned by the shift, add one to relevant hour
                for hr in range(cells[day][shift].hour, cells[day][shift + 1].hour):
                    hours[hr] += 1
        
        
        if cells[day][8] != 0:
            if cells[day][8].hour != 0:

                # If night shift exists...
                if cells[day][7] != 0:
                    for hr in range(cells[day][7].hour, cells[day][7].hour + cells[day][8].hour):
                        hours[hr % 24] += 1
                
                # Else if evening shift exists...
                elif cells[day][5] != 0:
                    for hr in range(cells[day][5].hour, cells[day][5].hour + cells[day][8].hour):
                        hours[hr % 24] += 1
                # If No evening/night shift, assume get-out starts at 10pm
                else:
                    for hr in range(22, 22 + cells[day][8].hour):
                        hours[hr % 24] += 1

                if debugHourCount:
                        print(f"GO: {cells[day][8].hour}", end = ' ')

        if debugHourCount:
            print()

""" Sum and return total hours counted. """
def sumHours(hours):
    return sum([hours[hr] for hr in hours])

""" Returns the absolute path to the timesheets to process. """
def getTimesheetDirPath():
    script_dir = os.path.dirname(__file__) 
    rel_path = timesheetsLocation
    return os.path.join(script_dir, rel_path)

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

def genHourDict():
    h = {}
    for i in range(24):
        h[i] = 0
    return h

indToDay = genIndToDayDict()
hours = genHourDict()

main()