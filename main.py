import os
import xlsxwriter
import time
import tkinter
import json

CONFIGFILE = "custom-config.json"
DELIMITER = ","
HEADERS = {"TC Number":-1, 
           "Risk - Total":-1, 
           "Region":-1, 
           "RWY (group)":-1,
           "Mile":-1,
           "Subdivision Name":-1,
           "Spur Mile":-1,
           "Spur Name":-1,
           "Date Inspected":-1,
           "Inspected By":-1,
           "Protection Type":-1}
FORCEHEADER = True #forces unmatched headers to have a column number.

def main():
    settings = ReadSettings() #check for settings file and load.
    print(settings)
    MainWindow = tkinter.Tk()
    MainWindow.title("GradeXConvertToXLSX")
    MainWindow.mainloop()
    
    ConvertToXLSX()
    
    
def ConvertToXLSX():
    writedir = "converted/"
    date = time.strftime("%d-%m-%y")
    workbooknum = 0
    
    for filename in os.listdir("."):
        if filename.endswith("csv"):
            if not os.path.isdir(writedir):
                os.makedirs(writedir)
            workbookname = writedir + "GradeXConv-" + date + "-" + str(workbooknum) + ".xlsx"
            workbook = xlsxwriter.Workbook(workbookname)
            worksheet = workbook.add_worksheet()
            wbheaderformat = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
            
            print("Reading " + filename)
            with open(filename, "r", encoding="utf-8-sig") as csvfile:
                lineat = 0
                for line in csvfile:
                    line = line.strip("\n").split(DELIMITER)
                    if lineat == 0: #header business
                        index = 0
                        for header in line:
                            if header in HEADERS:
                                HEADERS[header] = index
                            index += 1
                    else:
                        pass
                    
                    #write to CSV file
                    row = lineat
                    col = 0
                    for header in sorted(HEADERS):
                        if (FORCEHEADER and HEADERS[header] < 0): #check for unmatched headers.
                            if (row == 0):
                                worksheet.write(row, col, header, wbheaderformat)
                            col += 1
                            continue
                        elif HEADERS[header] < 0:
                            continue
                        
                        if row == 0:
                            worksheet.write(row, col, line[HEADERS[header]],wbheaderformat)
                        else:
                            worksheet.write(row, col, line[HEADERS[header]])
                        col += 1
            
                        
                        
                        
                    lineat += 1
            workbooknum += 1
            workbook.close()

def WriteSettings():
    with open(CONFIGFILE, "w") as configfile:
        settings = {"DELIMITER":DELIMITER, "HEADERS":HEADERS,"FORCEHEADER":FORCEHEADER}
        json.dump(settings, configfile, sort_keys=True, indent = 4)
        
def ReadSettings():
    # read settings file and update if needed.
    if os.path.isfile(CONFIGFILE):
        with open(CONFIGFILE, "r") as configfile:
            try:
                settings = json.load(configfile)
                if "DELIMITER" in settings:
                    DELIMITER = settings["DELIMITER"]
                if "HEADERS" in settings:
                    HEADERS = settings["HEADERS"]
                if "FORCEHEADER" in settings:
                    FORCEHEADER = settings["FORCEHEADER"]
            except:
                raise
        globals().update(settings)
        return True
    return False
            
if __name__ == "__main__":
    try:
        main()
        
    except:
        print("Unhandled Error.")
        raise