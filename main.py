'''
Created on Nov 18, 2016

@author: Matthew Muresan
''' 
import traceback
import os
import xlsxwriter
import datetime
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import tkinter.filedialog as tkf
import json
import threading
import difflib
import csv
import collections
import copy
from tkinter.filedialog import askdirectory
from difflib import get_close_matches

CONFIGFILE = "custom-config-4.json"
DELIMITER = ","
QUOTECHAR = '"'
HEADERS = {"DATABASE": collections.OrderedDict(), "PASSIVE": collections.OrderedDict(), "AWS": collections.OrderedDict(), "WIS": collections.OrderedDict(), "WSS": collections.OrderedDict(), "RANK": collections.OrderedDict()}
HEADERS["PASSIVE"]["Location Original ID"]= {"pos": -1, "alias":"TC Crossing ID" }
HEADERS["PASSIVE"]["Risk - Total"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Region"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Subdivision Mile Point"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Subdivision"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Spur Mile Point"]= {"pos": -1, "alias":False}
HEADERS["PASSIVE"]["Spur"]= {"pos": -1, "alias":False }
HEADERS["PASSIVE"]["Last  Inspection Date"]= {"pos": -1, "alias":"Last InspectionDate" }
HEADERS["PASSIVE"]["Last  Inspection By"]= {"pos": -1, "alias":"Last InspectionBy" }
HEADERS["PASSIVE"]["Passive Protection"]= {"pos": -1, "alias": False }
HEADERS["AWS"]["Location Original ID"]= {"pos": -1, "alias":"TC Crossing ID" }
HEADERS["AWS"]["Risk - Total"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Region"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Subdivision Mile Point"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Subdivision"]= {"pos": -1, "alias": False }
HEADERS["AWS"]["Spur Mile Point"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Spur"]= {"pos": -1, "alias":False }
HEADERS["AWS"]["Last  Inspection Date"]= {"pos": -1, "alias":"Last Inspection Date" }
HEADERS["AWS"]["Last  Inspection By"]= {"pos": -1, "alias":"Last Inspection By" }
HEADERS["AWS"]["AWS Protection"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Location Original ID"]= {"pos": -1, "alias":"Transport Canada WIS ID" }
HEADERS["WIS"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Subdivision Mile Point"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Subdivision"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Last  Inspection Date"]= {"pos": -1, "alias":"Last Inspection Date" }
HEADERS["WIS"]["Last  Inspection By"]= {"pos": -1, "alias":"Last Inspection By" }
HEADERS["WIS"]["Province"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Region"]= {"pos": -1, "alias":False }
HEADERS["WIS"]["Type"]= {"pos": -1, "alias": False}
HEADERS["WIS"]["Road Name Highway #"]= {"pos": -1, "alias":False}
HEADERS["WSS"]["Location Original ID"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Subdivision Mile Point"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Subdivision"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Last  Inspection Date"]= {"pos": -1, "alias":"Last Inspection Date" }
HEADERS["WSS"]["Last  Inspection By"]= {"pos": -1, "alias":"Last Inspection By" }
HEADERS["WSS"]["Province"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Region"]= {"pos": -1, "alias":False }
HEADERS["WSS"]["Type"]= {"pos": -1, "alias": False}
HEADERS["WSS"]["Road Name Highway #"]= {"pos": -1, "alias":False}
'''
#HEADERS["RANK"]["Rank"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["TC  Crossing ID"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Risk - Total"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Region"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Province"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Access"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Jurisdiction"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Subdivision Mile Point"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Subdivision"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Spur Mile Point"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Spur"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Latitude"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Longitude"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Road  Authority  #1"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["AWS Protection"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Passive Protection"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Accident #"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Total Injuries"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Fatalities"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Total Trains Daily"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Latest Train Traffic: Overall Maximum Speed (MPH) for Rail Approach from Left"]= {"pos": -1, "alias":"Train Max Speed (mph)" }
HEADERS["RANK"]["Tracks"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Latest Vehicle Traffic: # Vehicles Per Day"]= {"pos": -1, "alias":"Vehicles Daily"}
HEADERS["RANK"]["Max Road Speed"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["Lanes"]= {"pos": -1, "alias":False }
HEADERS["RANK"]["IsUrban"]= {"pos": -1, "alias":False }
'''
HEADERS["DATABASE"]["Risk - Total"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["TC  Crossing ID"]= {"pos": -1, "alias":"TC Number" }
HEADERS["DATABASE"]["Railway"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Region"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Province"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Access"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Jurisdiction"]= {"pos": -1, "alias":"Regulator" }
HEADERS["DATABASE"]["Subdivision Mile Point"]= {"pos": -1, "alias":"Mile" }
HEADERS["DATABASE"]["Subdivision"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Spur Mile Point"]= {"pos": -1, "alias":"Spur Mile" }
HEADERS["DATABASE"]["Spur"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Road Name Highway #"]= {"pos": -1, "alias":"Location" }
HEADERS["DATABASE"]["Latitude"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Longitude"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Road  Authority  #1"]= {"pos": -1, "alias":"Road Authority" }
HEADERS["DATABASE"]["AWS Protection"]= {"pos": -1, "alias":"Protection" }
#HEADERS["DATABASE"]["Passive Protection"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Accident #"]= {"pos": -1, "alias":"Accident" }
HEADERS["DATABASE"]["Fatalities"]= {"pos": -1, "alias":"Fatality" }
HEADERS["DATABASE"]["Total Injuries"]= {"pos": -1, "alias":"Injury" }
HEADERS["DATABASE"]["Total Trains Daily"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Latest Vehicle Traffic: # Vehicles Per Day"]= {"pos": -1, "alias":"Vehicles Daily"}
HEADERS["DATABASE"]["Latest Train Traffic: Overall Maximum Speed (MPH) for Rail Approach from Left"]= {"pos": -1, "alias":"Train Max Speed (mph)" }
HEADERS["DATABASE"]["Max Road Speed"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Lanes"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["Tracks"]= {"pos": -1, "alias":False }
HEADERS["DATABASE"]["IsUrban"]= {"pos": -1, "alias":False }


FILENAMEPREPEND = ""

SUMMATIONCOL = "Railway"
SUMMATIONS = ["Canadian National Railway", "Canadian Pacific Railway"]
SUMMATIONDO = True
NOSUMMATION = ["RANK", "EngGradeX", "First 500 Crossings English"]

REMAPDICT = {}
REMAPDICT["Province"] = {"Ont.":"ON", "Man.":"MB", "B.C.":"BC", "Que.":"QC", "N.B.":"NB", "N.S.":"NS", 
                 "P.E.I.":"PE", "Sask.":"SK", "Man.":"MN", "Nun.":"NU", "N.W.T.":"NT", 
                 "Nfld.":"NL", "Yuk.":"YT", "Yukn.":"YT", "Alta.": "AB"}
REMAPDICT["Jurisdiction"] = {"Federal": "F", "Provincial":"P"}
REMAPDICT["IsUrban"] = {"0": "N", "1":"Y"}

#PROVINCEPOSTALCONV = True # converts provinces to postal abbreviations
SORTRANK = "Risk - Total" #forces the "Rank" file to be sorted by rank if defined
FORCEHEADER = True #forces unmatched headers to have a column number.
RUNNING = False
OUTSIDEKILL = False

SPLITFILEON = "Region"
DONOTSPLITFILE = ["DATABASE"]#, "RANK"] #files that won't be split

ALLOWPARTIALFILE = True

FUZZYMATCHING = True #will attempt to match missing headers.
FUZZYLIMIT = 0.75 #the fuzzy level

FILETYPES = ["AWS", "PASSIVE", "WIS", "WSS", "DATABASE", "write"]#, "RANK"]

MainWindow = tk.Tk()
MainWindow.title("GradeXConvertToXLSX")
MainWindow.protocol('WM_DELETE_WINDOW', lambda: CloseProgram(MainWindow, None))

WARNINGS = False


class XLSWorkbook():
    def __init__(self, file, name, firsTabName=None):
        self.XLSXfile = file
        self.name = name
        self.worksheets = {}
        if firsTabName != None:
            self.AddWorksheet(firsTabName)
        
    #create a worksheet with name, if the header source in HEADERS is different than the name, define in headername
    def AddWorksheet(self, name, headername=False):
        self.worksheets[name] = XLSWorksheet(name, self.XLSXfile.add_worksheet(name), 
             self.XLSXfile.add_format({'bold': False, 'font_color': 'black', 'bg_color': '#D9D9D9'}), headername)
        
    def WriteLine(self, headers, line, tab):
        #if PROVINCEPOSTALCONV and "Province" in headers and headers["Province"]["pos"] > 0: #convert the province labesl
        #    if line[headers["Province"]["pos"]] in PROVINCEREMAP:
        #         line[headers["Province"]["pos"]] = PROVINCEREMAP[line[headers["Province"]["pos"]]]
        for header in headers: #iterate through the headers and assemble what needs to be written
            if header in REMAPDICT: #there is a translation/remapping that is needed
                if line[headers[header]["pos"]] in REMAPDICT[header]: #check if the remap for this case is defined:
                    line[headers[header]["pos"]] = REMAPDICT[header][line[headers[header]["pos"]]] #replace with remap value
            if not tab in self.worksheets: #if the tab doesn't exist make it.
                self.AddWorksheet(tab)
            if not headers[header]["pos"] < 0:
                self.worksheets[tab].writeCell(line[headers[header]["pos"]])
            
        self.worksheets[tab].nextRow()
            
    def close(self):
        self.XLSXfile.close()
    
class XLSWorksheet():
    def __init__(self, name, worksheet, headerformat, headername=False):
        self.name = name
        if headername == False:
            headername = name #create a default tabname with the nam
        self.worksheet = worksheet
        self.atRow = 0
        self.atCol = 0
        self.maxCol = 0
        self.summationRWYcol = None
        self.wbheaderformat = headerformat
        
        #write header
        for header in HEADERS[headername]:
            if HEADERS[headername][header]["alias"] != False:
                self.worksheet.write(self.atRow, self.atCol, HEADERS[headername][header]["alias"], self.wbheaderformat)
            else:
                self.worksheet.write(self.atRow, self.atCol, header, self.wbheaderformat)
            if header == SUMMATIONCOL:
                self.summationRWYcol = chr(65 + self.atCol) #store the letter of the column
            self.atCol += 1
        
        self.maxCol = self.atCol
        
        if SUMMATIONDO and not self.name in NOSUMMATION: self.writeSummation()        
        
        self.nextRow()
    
    def writeSummation(self):
        Trow = 0
        Tcol = self.maxCol + 2 #skip a column
        
        self.worksheet.write(Trow, Tcol, "RWY", self.wbheaderformat)
        self.worksheet.write(Trow, Tcol + 1, self.name, self.wbheaderformat)
        
        for summation in SUMMATIONS:
            Trow += 1
            sumfor = "=COUNTIF(" + self.summationRWYcol + ":" + self.summationRWYcol + "," + chr(65 + Tcol) + str(Trow + 1) + ")"
            self.worksheet.write(Trow, Tcol, summation)
            self.worksheet.write(Trow, Tcol + 1, sumfor)
            
        Trow += 1
        sumfor = "= " + (chr(65 + Tcol + 1) + str(Trow + 2 + 1) + " - (" + chr(65 + Tcol + 1) + str(Trow - 2 + 1) + " + " 
                  + chr(65 + Tcol + 1) + str(Trow - 1 + 1) + ")") 
        self.worksheet.write(Trow, Tcol, "Other")
        self.worksheet.write(Trow, Tcol + 1, sumfor)
        
        Trow += 2 #skip two rows to write the GT
        sumfor = "=COUNTA(" + self.summationRWYcol + ":" + self.summationRWYcol + ") - 1"
        self.worksheet.write(Trow, Tcol, "Total")
        self.worksheet.write(Trow, Tcol + 1, sumfor)
    def nextRow(self):
        self.atRow += 1
        self.atCol = 0 #advance and reset the rows.
        
    def writeCell(self, item):
        try:
            float(item)
            item = float(item)
            self.worksheet.write_number(self.atRow, self.atCol, item)
        except ValueError:
            self.worksheet.write(self.atRow, self.atCol, item)
            
        self.atCol += 1
    
    def is_number(self, s):
        try:
            float(s)
            return True
        except ValueError:
            return False

def main():
    Files = {}
    for name in FILETYPES:
        Files[name] = False
    
    settings = ReadSettings() #check for settings file and load.
    #widget def
    config = ttk.Button(MainWindow, text="Configure", command=ShowConfig)
    textlbl = ttk.Label(MainWindow, text='Application Messages'
                            ,width=75, wraplength=550, justify=tk.LEFT, padding=(12,12,12,12))
    messagelist = tk.Listbox(MainWindow, height=8, width=100)
    
    mode_opt = tk.StringVar()
    mode_opt.set("S")
    
    ok = ttk.Button(MainWindow, text="Run", command=lambda: RunApplication(messagelist, MainWindow, Files, mode_opt))
    close = ttk.Button(MainWindow, text="Close", command=lambda: CloseProgram(MainWindow, messagelist))
    
    brwslocAWS = ttk.Entry(MainWindow)
    brwslocAWS.insert(0, "Please Load a AWS GradeX Output File")
    brwslocAWS.configure(state='disabled')
    brwslocPAS = ttk.Entry(MainWindow)
    brwslocPAS.insert(0, "Please Load a GradeX Output File for PASSIVE Crossings")
    brwslocPAS.configure(state='disabled')
    brwslocWIS = ttk.Entry(MainWindow)
    brwslocWIS.insert(0, "Please Load a WIS GradeX Output File")
    brwslocWIS.configure(state='disabled')
    brwslocWSS = ttk.Entry(MainWindow)
    brwslocWSS.insert(0, "Please Load a WSS GradeX Output File")
    brwslocWSS.configure(state='disabled')
    brwslocLST = ttk.Entry(MainWindow)
    #brwslocLST.insert(0, "Please Load a List Ranking File")
    #brwslocLST.configure(state='disabled')
    brwslocCOM = ttk.Entry(MainWindow)
    brwslocCOM.insert(0, "Provide a complete database file")
    brwslocCOM.configure(state='disabled')
    
    mode_opt_l = ttk.Label(MainWindow, text="Select how the input is provided:", justify=tk.CENTER, padding=(12,0,0,0))
    
    
    brwslocW = ttk.Entry(MainWindow)
    brwslocW.insert(0, "Please Choose a Directory to Save all Output to")
    brwslocW.configure(state='disabled')
    readfileBAWS = ttk.Button(MainWindow, text="AWS...", command=lambda: askFile("AWS", Files, brwslocAWS, MainWindow))
    readfileBPAS = ttk.Button(MainWindow, text="PASSIVE...", command=lambda: askFile("PASSIVE", Files, brwslocPAS, MainWindow))
    readfileBWIS = ttk.Button(MainWindow, text="WIS...", command=lambda: askFile("WIS", Files, brwslocWIS, MainWindow))
    readfileBWSS = ttk.Button(MainWindow, text="WSS...", command=lambda: askFile("WSS", Files, brwslocWSS, MainWindow))
    #readfileBLST = ttk.Button(MainWindow, text="Rank...", command=lambda: askFile("RANK", Files, brwslocLST, MainWindow))
    readfileBCOM = ttk.Button(MainWindow, text="Database...", command=lambda: askFile("DATABASE", Files, brwslocCOM, MainWindow))
    writefileB = ttk.Button(MainWindow, text="Save to...", command=lambda: askFile("write", Files, brwslocW, MainWindow))
    
    items = {"readfileBAWS":readfileBAWS, "brwslocAWS":brwslocAWS,
             "readfileBPAS":readfileBPAS, "brwslocPAS":brwslocPAS,
             "readfileBWIS":readfileBWIS, "brwslocWIS":brwslocWIS,
             "readfileBWSS":readfileBWSS, "brwslocWSS":brwslocWSS,
             #"readfileBLST":readfileBLST, "brwslocLST":brwslocLST,
             "readfileBCOM":readfileBCOM, "brwslocCOM":brwslocCOM}
             
    mode_laba = ttk.Radiobutton(MainWindow, text="Rank File (For Web)", variable=mode_opt, value="S", command=lambda: re_select(items, "S"))
    mode_labb = ttk.Radiobutton(MainWindow, text="Other Files", variable=mode_opt, value="M", command=lambda: re_select(items, "M"))
    
    
    notelabel = ttk.Label(MainWindow, text="Note: Omitted files will not \nbe included in the final export.", justify=tk.CENTER, padding=(12,0,0,0))
    #widget layout
    #textlbl.grid(row=0, column=1, columnspan=3)
    messagelist.grid(row=2, column=1, columnspan=3, sticky=(tk.N, tk.S, tk.E, tk.W), pady=20)
    messagelist.insert(tk.END, "GradeX Output Converter")
    if not settings: messagelist.insert(tk.END, "  Select a File to Convert...")
    else: messagelist.insert(tk.END, "  Custom settings loaded! Select a file to convert...")
    
    ok.grid(row=99, column=1, pady=5)
    config.grid(row=98, column=1, pady=5)
    close.grid(row=99, column=3, pady=5)
    
    notelabel.grid(row=85, column=1, pady=0, columnspan = 1, rowspan = 3, sticky=(tk.W))
    mode_opt_l.grid(row=88, column=1, pady=0, columnspan = 1, rowspan = 1, sticky=(tk.W))
    mode_laba.grid(row=89, column=1, pady=0, padx=25, columnspan = 1, rowspan = 1, sticky=(tk.W))
    mode_labb.grid(row=90, column=1, pady=0, padx=25, columnspan = 1, rowspan = 1, sticky=(tk.W))
    
    writefileB.grid(row=91, column=1, pady=20, sticky=(tk.E))
    brwslocW.grid(row=91, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBAWS.grid(row=86, column=1, pady=0, sticky=(tk.E))
    brwslocAWS.grid(row=86, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBPAS.grid(row=87, column=1, pady=0, sticky=(tk.E))
    brwslocPAS.grid(row=87, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBWIS.grid(row=88, column=1, pady=0, sticky=(tk.E))
    brwslocWIS.grid(row=88, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBWSS.grid(row=89, column=1, pady=0, sticky=(tk.E))
    brwslocWSS.grid(row=89, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    #readfileBLST.grid(row=85, column=1, pady=0, sticky=(tk.E))
    #brwslocLST.grid(row=85, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBCOM.grid(row=85, column=1, rowspan=5, pady=0, sticky=(tk.E))
    brwslocCOM.grid(row=85, column=2, rowspan=5, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    re_select(items, "S")
    
    MainWindow.resizable(width=False, height=False)
    MainWindow.mainloop()
    

class error_catch:
    def __init__(self, function):
        self.function = function
    
    def __call__(self, *args):
        try:
            return self.function(*args)
        except Exception as e:
            s = str(e)
            #print("Unhandled Error.")
            messagebox.showerror(title="Unhandled Error.", message=s)

@error_catch
def RunApplication(updatebox, window, Files, mode_opt):
    def callback():
        ConvertToXLSX(updatebox, window, Files, mode_opt)
        
    if not RUNNING:
        updatebox.insert(tk.END, "    Converting Files...")
        updatebox.yview(tk.END)    
        window.update()
        thread = threading.Thread(target=callback)
        thread.start()
            
@error_catch
def ShowConfig():
    ConfigWindow = tk.Toplevel()
    ConfigWindow.title("Configure GradeXConvertToXLSX")
    
    headerssting = ""
    firstheader = True
    for key in sorted(HEADERS):
        if firstheader:
           firstheader = False
        else:
            headerssting += "," 
        headerssting += key
    forcedhead = tk.IntVar()   
    forcedhead.set(FORCEHEADER) 

    headerstolbl = ttk.Label(ConfigWindow, text="Press 'Export' to export the existing settings to a configurable file. \n Configuration files are automatically loaded when you re-load the program. \n Press 'Load' to refresh the settings if you've made changes to the custom configuration file. \nRestore defaults by deleting the custom-config.json file in the program's directory", justify=tk.CENTER, padding=(12,24,12,0))
    ok = ttk.Button(ConfigWindow, text="Export Settings to File", command=lambda: WriteSettings(ConfigWindow))
    load = ttk.Button(ConfigWindow, text="Load New Settings File", command=lambda: ReadSettings(ConfigWindow, True))
    close = ttk.Button(ConfigWindow, text="Close", command=ConfigWindow.destroy)
    
    
    headerstolbl.grid(row=3, column=1, columnspan=2)
    
    ok.grid(row=54, column=1, pady=20)
    close.grid(row=55, column=2, pady=10)
    load.grid(row = 54, column = 2, pady = 10)
    
    ConfigWindow.mainloop()

def CloseProgram(window, updatebox):
    if RUNNING:
        global OUTSIDEKILL
        OUTSIDEKILL = True
        
        if updatebox != None:
            updatebox.insert(tk.END, "    Abortring run...")
            updatebox.yview(tk.END)    
            window.update()
        
    else:
        window.destroy()

def _CheckFiles(files, filename):
    if not filename in files:
        return -2 #the file is not being considered
    if files[filename] == False or files[filename] == None: #check that files exist
        return -1
    if filename not in HEADERS and filename != "write": #check that headers are defined for file.
        return -2
    return 1

def _ProcessFiles(updatebox, window, files, name, workbooks, now):
    filename = files[name].name
    global OUTSIDEKILL

    
    lineat = 0
    lines = []
    with open(filename, "r") as csvfile:
        csvrd = csv.reader(csvfile, delimiter=DELIMITER, quotechar=QUOTECHAR)
        for line in csvrd:
            #line = line.strip("\n").split(DELIMITER)
            if lineat == 0: #header business
                index = 0
                #figure out the location of all the headers.
                for header in line:
                    if header in HEADERS[name]:
                        HEADERS[name][header]["pos"] = index
                    index += 1
                    
                for header in sorted(HEADERS[name]): #check if all headers matched
                    if HEADERS[name][header]["pos"] < 0:
                        updateboxtext = "               WARNING!! Column \"" + header + "\" not found in file " + filename + "! " 
                        if header == SPLITFILEON: 
                            updateboxtext += SPLITFILEON 
                            updateboxtext+= " is required to build the files! Unable to process file!"
                        WriteWarnings(updateboxtext)
                        updatebox.insert(tk.END, updateboxtext)
                        updatebox.yview(tk.END)
                        window.update()
                        if FUZZYMATCHING: #If fuzzy matching enabled, try to match to something close
                            fuzzymatch = get_close_matches(header, line, n=1, cutoff=FUZZYLIMIT)
                            if len(fuzzymatch) > 0:
                                HEADERS[name][header]["pos"] = line.index(fuzzymatch[0])
                                updateboxtext = "                 " + header + " column not found! Using closest match " + fuzzymatch[0]
                                WriteWarnings(updateboxtext)
                                updatebox.insert(tk.END, updateboxtext)
                                updatebox.yview(tk.END)
                                window.update()
                        if header == SPLITFILEON : return #important header, terminate
                        continue
                lineat += 1
                continue #read the next line of the file.
            #write to CSV file            
            
            if name != "DATABASE" and SORTRANK == False: #Legacy multiple file mode. Write each line individually without sorting
                workbookname = line[HEADERS[name][SPLITFILEON]["pos"]]
                if name in DONOTSPLITFILE:
                    workbookname = name #don't split these files. Give them the name "name"
                if workbookname == "" or workbookname == None:
                    updateboxtext = "               Row #" + str(lineat) + " has no Region!"
                    WriteWarnings(updateboxtext)
                    updatebox.insert(tk.END, updateboxtext)
                    updatebox.yview(tk.END)
                    workbookname = "No Region"
                
                if not workbookname in workbooks:
                    if not os.path.exists(files["write"] + "/" + now):
                        os.makedirs(files["write"] + "/" + now)
                    wbn = files["write"] + "/" + now + "/" + FILENAMEPREPEND + workbookname + ".xlsx"
                    wbf = xlsxwriter.Workbook(wbn)
                    workbooks[workbookname] = XLSWorkbook(wbf, workbookname, name) 
    
                workbooks[workbookname].WriteLine(HEADERS[name], line, name)
     
                lineat += 1
                if (lineat % 5000 == 0):
                    updateboxtext = "           Processed " + str(lineat) + " rows" 
                    updatebox.insert(tk.END, updateboxtext)
                    updatebox.yview(tk.END)    
                    window.update()
                if OUTSIDEKILL:
                    #kills the program if requested. This quits without saving the workbook.
                    OUTSIDEKILL = False
                    RUNNING = False
                    messagebox.showinfo(title="Aborted Run", message="Run aborted, output file not complete!")
                    return
            else:
                try:
                    line[HEADERS[name][SORTRANK]["pos"]] = float(line[HEADERS[name][SORTRANK]["pos"]])
                except:
                    line[HEADERS[name][SORTRANK]["pos"]] = 0
                lines.append(line)
    if SORTRANK != False and name == "DATABASE":
        HEADERS[name][SORTRANK]["alias"] = "Rank" #rename field to "Rank" as it is the rank sort.
        lineat = 0
        updateboxtext = "      Sorting based on " + str(SORTRANK) + " and formatting data..."
        updatebox.insert(tk.END, updateboxtext)
        updatebox.yview(tk.END) 
        window.update()
        
        lines.sort(key=lambda x: x[HEADERS[name][SORTRANK]["pos"]], reverse=True)
        workbookname = name
        tabname = "EngGradeX"
        tabname500 = "First 500 Crossings English"
        if not workbookname in workbooks:
            if not os.path.exists(files["write"] + "/" + now):
                os.makedirs(files["write"] + "/" + now)
            wbn = files["write"] + "/" + now + "/" + FILENAMEPREPEND + workbookname + ".xlsx"
            wbf = xlsxwriter.Workbook(wbn)
            workbooks[workbookname] = XLSWorkbook(wbf, workbookname) 
            workbooks[workbookname].AddWorksheet(tabname, "DATABASE")
            workbooks[workbookname].AddWorksheet(tabname500, "DATABASE")
        for line in lines:
            line[HEADERS[name][SORTRANK]["pos"]] = lineat + 1
            workbooks[workbookname].WriteLine(HEADERS[name], line, tabname)
            if lineat < 500:
                workbooks[workbookname].WriteLine(HEADERS[name], line, tabname500)
            
            lineat += 1
            
            
            if (lineat % 5000 == 0):
                updateboxtext = "           Processed " + str(lineat) + " rows" 
                updatebox.insert(tk.END, updateboxtext)
                updatebox.yview(tk.END)    
                window.update()
    
def ConvertToXLSX(updatebox, window, files, mode_opt):
    global RUNNING
    global OUTSIDEKILL
    global FILETYPES
    global WARNINGS
    RUNNING = True
    workbooks = {} #where we store the active workbooks as we write too them.
    #cleanup unneeded files
    filetypes = []
    if mode_opt.get() == "M":
        for name in FILETYPES:
            if name != "DATABASE":
                filetypes.append(name)
    else:
        filetypes.append("DATABASE")
        filetypes.append("write")
        
    try:            
        for name in filetypes:
            i = _CheckFiles(files, name)
            if i != 1: #error detected
                if i == -1:
                    if ALLOWPARTIALFILE and name != "write":
                        updateboxtext = "WARNING! File " + name + " not included!"
                        WriteWarnings(updateboxtext)
                        updatebox.insert(tk.END, updateboxtext)
                        updatebox.yview(tk.END)  
                        continue
                    else:
                        updateboxtext = "File '" + name + "' not selected. No data processed!"
                elif i == -2:
                    updateboxtext = "Configuration files are not correct. No data processed!"
                else:
                    updateboxtext = "General error while checking validity of files, No data processed!"
                updatebox.insert(tk.END, updateboxtext)
                updatebox.yview(tk.END)    
                window.update()
                return #don't do anything if files invalid
                
            
        now = datetime.datetime.now()
        now = now.strftime("OUTPUT %Y-%m-%d %H%M %S")
        for name in filetypes: #iterate through the files
            if name == "write": continue
            if files[name] == False: continue #skip if missing
            filename = files[name].name
            files[name].close()
            updateboxtext = "      Reading File " + filename 
            updatebox.insert(tk.END, updateboxtext)
            updatebox.yview(tk.END)    
            window.update()
            _ProcessFiles(updatebox, window, files, name, workbooks, now) #load all files into memory
            
        
        
        updateboxtext = "Please wait, files being written to " + files["write"]
        updatebox.insert(tk.END, updateboxtext)
        updatebox.yview(tk.END)    
        window.update()
    except Exception as e:
        ShowErrors(str(e))
        with open("error.log", 'w') as errfile:
            str(traceback.print_exc(file=errfile))
        raise
    finally:
        #close all the workbooks (if any) open
        RUNNING = False
        if WARNINGS:
            messagebox.showinfo(title="An Exception Occured!", message='Some warnings were raised! Please review the message log or see "warn.log" for mor information')
            WARNINGS = False
        if len(workbooks) > 0:
            for wbname in workbooks:
                workbooks[wbname].close()
            
            updateboxtext = "All Files Written"
            updatebox.insert(tk.END, updateboxtext)
            updatebox.yview(tk.END)    
            window.update()
            
def WriteWarnings(text):
    global WARNINGS
    WARNINGS = True
    with open("warn.log", "a") as warnf:
        warnf.write("at " + str(datetime.datetime.now()) + "\n" + text)
        warnf.write("\n\n")        

def ShowErrors(errorT):
    messagebox.showinfo(title="An Exception Occured!", message='An error occured while processing files... the error text was: \n' + errorT + "\n\nSee error.log for more information")
    

def WriteSettings(ConfigWindow):
    global ALLOWPARTIALFILE
    global DELIMITER
    global QUOTECHAR
    global SUMMATIONCOL
    global SUMMATIONDO
    global SPLITFILEON
    global DONOTSPLITFILE
    global NOSUMMATION
    global REMAPDICT
    global SORTRANK
    #global PROVINCEREMAP
    #global PROVINCEPOSTALCONV
    global HEADERS
    global FORCEHEADER
    global FUZZYMATCHING
    global FUZZYLIMIT
    global FILENAMEPREPEND
    global PROVINCEPREMAP
    global SUMMATIONS
    
    comments = {"1":"This file contains a list of all the configurable variables. These instructions briefly highlight what each field does. Delete this file to restore defaults. This is a JSON file, you must write correct JSON code for it to be parsed. See http://www.json.org/ for more information. Further help with configuration can also be obtained by contacting mimuresa@uwaterloo.ca. See also https://github.com/jargon777/GradeXConvert for code.",
                "ALLOWPARTIALFILE": "Boolean. Whether or not the program will run if some of the files are missing. Default is false, meaning AWS, PASSIVE, WIS and WSS CSV inputs must be provided.",
                "DELIMTER": "The delimiter used to separate values in the input CSV file.",
                "DONOTSPLITFILE": "Files that will not be split according to SPLIFILEON",
                "FILENAMEPREPEND": "This text is prepended to the filenames of the workbooks written, e.g. if it is set to 'test' then the AWS workbook will be 'testAWS.xlsx'. These names are not santized, and invalid characters will throw a file name error on execution.",
                "FORCEHEADER": "Boolean. If a suitable match for the column can't be found, then should the column be kept?",
                "FUZZYMATCHING": "Boolean. If set to true, the program will attempt to find a close match for the requested column. If set to false, exact matches are needed to column names. Fuzzy matching issues warnings to the screen if used.",
                "FUZZYLIMIT": "Integer between 0 and 1, configures the limit of the fuzzy matching. 1 means more strict matching, 0 less strict.",
                "HEADERS": "The headers included in the input files. The overall format is a dictionary datatype, declared as HEADERS['TAB']['NAME IN INPUT'] = {'pos':1, 'alias': NAME IN OUTPUT}. In laymans terms, this setting has three levels. In the first level, each of the tags is used to generate separate tabs in the files. This cannot be changed in the setting file without throwing errors. The second level has the names of the columns from the source files that will be included in the written file. Be sure the column exists if you change these values. The third level has two keys, 'pos' and 'remap'. 'pos' is used when converting, and must take the form 'pos': -1. 'remap' controls what is written to the ouput file. Existing headers can be renamed. The order of entries determines the order in the output.",
                "NOSUMMATION": "List. If SUMMATIONDO is set to True, then these tabs will NOT get a summation box.",
                #"PROVINCEPOSTALCONV": "Boolean. Whether or not the program will try to convert the province names from non-standard abbreviations to postal abbreviations",
                #"PROVINCEPREMAP": "The names that are converted, in format 'Abbreviation : Postal Abbreviation'. Values matching 'Abbreviuation' will be converted to 'Postal Abbreviation'",
                "QUOTECHAR": "The quoting used to group together values that the delimiter in them in the input CSV file.",
                "REMAPDICT": "Each defined key will have its values remapped based on the provided dictionary. For example, if 'Province' is defined with dictionary {'Ont.': 'ON'} 'Ont.' in the input will be replaced with 'ON' in the output.",
                "SORTRANK": "If not equal to False, will sort the outputted rank file when operating in single file mode by the given column",
                "SUMMATIONCOL": "Name of the column to use when generating the RWY sums",
                "SUMMATIONDO": "Boolean. Whether or not the summation statistics are written to the excel file",
                "SPLITFILEON": "String, the column that is used to separate the inputs into different files. Default is to create region-based files.",
                "SUMMATIONS": "List of the names included in the summary field."
                }
    
    for key in HEADERS:
        for item in HEADERS[key]:
            HEADERS[key][item]["pos"] = -1
    
    with open(CONFIGFILE, "w") as configfile:
        settings = {"DELIMITER":DELIMITER, "QUOTECHAR":QUOTECHAR, "HEADERS":HEADERS,"FORCEHEADER":FORCEHEADER,
                    "FUZZYMATCHING":FUZZYMATCHING, "SUMMATIONCOL":SUMMATIONCOL, "REMAPDICT":REMAPDICT, 
                    "SORTRANK":SORTRANK, "FUZZYLIMIT":FUZZYLIMIT, "SUMMATIONDO":SUMMATIONDO,
                    "NOSUMMATION":NOSUMMATION, "FILENAMEPREPEND":FILENAMEPREPEND, "ALLOWPARTIALFILE":ALLOWPARTIALFILE, 
                    "SPLITFILEON":SPLITFILEON, "SUMMATIONS":SUMMATIONS, "DONOTSPLITFILE":DONOTSPLITFILE,  "1-INSTRUCTIONS":comments}
        json.dump(settings, configfile, sort_keys=True, indent = 4)
    
    ConfigWindow.destroy()
    messagebox.showinfo(title="Configurations Saved", message='Configuration File saved as ' + CONFIGFILE)
    

def re_select(items, dir):
    if dir == "S":
        for item_key in items:
            if item_key == "readfileBCOM" or item_key == "brwslocCOM":
                items[item_key].grid()
            else: 
                items[item_key].grid_remove()
    else:
        for item_key in items:
            if item_key != "readfileBCOM" and item_key != "brwslocCOM":
                items[item_key].grid()
            else: 
                items[item_key].grid_remove()
   
def askFile(key, fileDict, label, window):
    global FILETYPES
    #Calls a dialog box that asks the user to navigate to a folder to save localconfig.
    if key == "write":
        #file = tkf.asksaveasfile("w", defaultextension=".xlsx", filetypes =(("Microsoft Excel Table", "*.xlsx"),("All Files","*.*")))
        file = tkf.askdirectory()
    elif key in FILETYPES:
        file = tkf.askopenfile("r", defaultextension=".csv", filetypes =(("GradeX Output", "*.csv"),("All Files","*.*")))
    if file != False and file != None:
        fileDict[key] = file
        fname = file.name if key != "write" else file
        label.configure(state='normal')
        label.delete(0, tk.END)
        label.insert(0, fname)
        label.configure(state='disabled')
        window.update_idletasks()
        
def ReadSettings(ConfigWindow=None, Message=False):
    # read settings file and update if needed.
    if os.path.isfile(CONFIGFILE):
        with open(CONFIGFILE, "r") as configfile:
            try:
                settings = json.load(configfile)
                if "DELIMITER" in settings:
                    DELIMITER = settings["DELIMITER"]
                if "QUOTECHAR" in settings:
                    QUOTECHAR = settings["QUOTECHAR"]
                if "HEADERS" in settings:
                    HEADERS = settings["HEADERS"]
                if "FORCEHEADER" in settings:
                    FORCEHEADER = settings["FORCEHEADER"]
                if "FUZZYMATCHING" in settings:
                    FUZZYMATCHING = settings["FUZZYMATCHING"]
                if "FUZZYLIMIT" in settings:
                    FUZZYLIMIT = settings["FUZZYLIMIT"]
                if "SUMMATIONCOL" in settings:
                    SUMMATIONCOL = settings["SUMMATIONCOL"]
                if "REMAPDICT" in settings:
                    REMAPDICT = settings["REMAPDICT"]
                if "SORTRANK" in settings:
                    SORTRANK = settings["SORTRANK"]
                if "SUMMATIONCOL" in settings:
                    SUMMATIONCOL = settings["SUMMATIONCOL"]
                if "SUMMATIONDO" in settings:
                    SUMMATIONDO = settings["SUMMATIONDO"]
                if "NOSUMMATION" in settings:
                    NOSUMMATION = settings["NOSUMMATION"]
                if "SPLITFILEON" in settings:
                    SPLITFILEON = settings["SPLITFILEON"]
                if "SUMMATIONS" in settings:
                    SUMMATIONS = settings["SUMMATIONS"]
                if "DONOTSPLITFILE" in settings:
                    DONOTSPLITFILE = settings["DONOTSPLITFILE"]
                if "FILENAMEPREPEND" in settings:
                    FILENAMEPREPEND = settings["FILENAMEPREPEND"]
                if "ALLOWPARTIALFILE" in settings:
                    ALLOWPARTIALFILE = settings["ALLOWPARTIALFILE"]
            except:
                raise
        globals().update(settings)
        if Message and ConfigWindow != None: 
            ConfigWindow.destroy()
            messagebox.showinfo(title="Configurations Saved", message='Configuration file ' + str(CONFIGFILE) + " loaded!")
        return True
    return False
    
if __name__ == "__main__":
    main()
        
