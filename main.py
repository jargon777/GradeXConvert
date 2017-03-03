'''
Created on Nov 18, 2016

@author: Matthew Muresan
''' 

import os
import xlsxwriter
import time
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.ttk as ttk
import tkinter.filedialog as tkf
import json
import threading
import difflib
import csv
import collections
from tkinter.filedialog import askdirectory
from difflib import get_close_matches

CONFIGFILE = "custom-config.json"
DELIMITER = ","
QUOTECHAR = '"'
'''
HEADERS = {"TC Number"] = -1, 
           "Risk - Total"] = -1, 
           "Region"] = -1, 
           "RWY (group)"] = -1,
           "Mile"] = -1,
           "Subdivision Name"] = -1,
           "Spur Mile"] = -1,
           "Spur Name"] = -1,
           "Date Inspected"] = -1,
           "Inspected By"] = -1,
           "Protection Type"] = -1}
'''
HEADERS = {"PASSIVE": collections.OrderedDict(), "AWS": collections.OrderedDict(), "WIS": collections.OrderedDict(), "WSS": collections.OrderedDict()}
HEADERS["PASSIVE"]["Location Original ID"] = -1
HEADERS["PASSIVE"]["Risk - Total"] = -1
HEADERS["PASSIVE"]["Region"] = -1
HEADERS["PASSIVE"]["Railway (group)"] = -1
HEADERS["PASSIVE"]["Subdivision Mile Point"] = -1
HEADERS["PASSIVE"]["Subdivision"] = -1
HEADERS["PASSIVE"]["Spur Mile Point"] = -1
HEADERS["PASSIVE"]["Spur"] = -1
HEADERS["PASSIVE"]["Last  Inspection Date"] = -1
HEADERS["PASSIVE"]["Last  Inspection By"] = -1
HEADERS["PASSIVE"]["Type"] = -1
HEADERS["AWS"]["Location Original ID"] = -1
HEADERS["AWS"]["Risk - Total"] = -1
HEADERS["AWS"]["Region"] = -1
HEADERS["AWS"]["Railway (group)"] = -1
HEADERS["AWS"]["Subdivision Mile Point"] = -1
HEADERS["AWS"]["Subdivision"] = -1
HEADERS["AWS"]["Spur Mile Point"] = -1
HEADERS["AWS"]["Spur"] = -1
HEADERS["AWS"]["Last  Inspection Date"] = -1
HEADERS["AWS"]["Last  Inspection By"] = -1
HEADERS["AWS"]["Type"] = -1
HEADERS["WIS"]["Location Original ID"] = -1
HEADERS["WIS"]["Railway (group)"] = -1
HEADERS["WIS"]["Subdivision Mile Point"] = -1
HEADERS["WIS"]["Subdivision"] = -1
HEADERS["WIS"]["Province"] = -1
HEADERS["WIS"]["Region"] = -1
HEADERS["WIS"]["Type"] = -1
HEADERS["WSS"]["Location Original ID"] = -1
HEADERS["WSS"]["Railway (group)"] = -1
HEADERS["WSS"]["Subdivision Mile Point"] = -1
HEADERS["WSS"]["Subdivision"] = -1
HEADERS["WSS"]["Province"] = -1
HEADERS["WSS"]["Region"] = -1
HEADERS["WSS"]["Type"] = -1

FORCEHEADER = True #forces unmatched headers to have a column number.
RUNNING = False
OUTSIDEKILL = False

FUZZYMATCHING = True #will attempt to match missing headers.

FILETYPES = ["AWS", "PASSIVE", "WIS", "WSS", "write"]

MainWindow = tk.Tk()
MainWindow.title("GradeXConvertToXLSX")
MainWindow.protocol('WM_DELETE_WINDOW', lambda: CloseProgram(MainWindow, None))


class XLSWorkbook():
    def __init__(self, file, name, firsTabName):
        self.XLSXfile = file
        self.name = name
        self.worksheets = {}
        
        self.AddWorksheet(firsTabName)
        
        
    def AddWorksheet(self, name):
        self.worksheets[name] = XLSWorksheet(name, self.XLSXfile.add_worksheet(name), 
             self.XLSXfile.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'}))
        
    def WriteLine(self, headers, line, tab):
        for header in headers: #iterate through the headers and assemble what needs to be written
            if not tab in self.worksheets: #if the tab doesn't exist make it.
                self.AddWorksheet(tab)
            if not headers[header] < 0:
                self.worksheets[tab].writeCell(line[headers[header]])
            
        self.worksheets[tab].nextRow()
            
    def close(self):
        self.XLSXfile.close()
    
class XLSWorksheet():
    def __init__(self, tabname, worksheet, headerformat):
        self.name = tabname
        self.worksheet = worksheet
        self.atRow = 0
        self.atCol = 0
        self.wbheaderformat = headerformat
        
        for header in HEADERS[tabname]:
            self.worksheet.write(self.atRow, self.atCol, header, self.wbheaderformat)
            self.atCol += 1
        
        self.nextRow()
    
    def nextRow(self):
        self.atRow += 1
        self.atCol = 0 #advance and reset the rows.
        
    def writeCell(self, item):
        self.worksheet.write(self.atRow, self.atCol, item)
        self.atCol += 1

def main():
    Files = {}
    for name in FILETYPES:
        Files[name] = False
    
    settings = ReadSettings() #check for settings file and load.
    #widget def
    config = ttk.Button(MainWindow, text="Configure", command=ShowConfig)
    textlbl = ttk.Label(MainWindow, text='Application Messages'
                            ,width=75, wraplength=550, justify=tk.LEFT, padding=(12,12,12,12))
    messagelist = tk.Listbox(MainWindow, height=3, width=80)
    
    ok = ttk.Button(MainWindow, text="Run", command=lambda: RunApplication(messagelist, MainWindow, Files))
    close = ttk.Button(MainWindow, text="Cancel", command=lambda: CloseProgram(MainWindow, messagelist))
    
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
    
    brwslocW = ttk.Entry(MainWindow)
    brwslocW.insert(0, "Please Choose a Directory to Save all Output to")
    brwslocW.configure(state='disabled')
    readfileBAWS = ttk.Button(MainWindow, text="AWS...", command=lambda: askFile("AWS", Files, brwslocAWS, MainWindow))
    readfileBPAS = ttk.Button(MainWindow, text="PASSIVE...", command=lambda: askFile("PASSIVE", Files, brwslocPAS, MainWindow))
    readfileBWIS = ttk.Button(MainWindow, text="WIS...", command=lambda: askFile("WIS", Files, brwslocWIS, MainWindow))
    readfileBWSS = ttk.Button(MainWindow, text="WSS...", command=lambda: askFile("WSS", Files, brwslocWSS, MainWindow))
    writefileB = ttk.Button(MainWindow, text="Save to...", command=lambda: askFile("write", Files, brwslocW, MainWindow))
    #widget layout
    #textlbl.grid(row=0, column=1, columnspan=3)
    messagelist.grid(row=2, column=1, columnspan=3, sticky=(tk.N, tk.S, tk.E, tk.W), pady=20)
    messagelist.insert(tk.END, "GradeX Output Converter")
    messagelist.insert(tk.END, "  Select a File to Convert...")
    
    ok.grid(row=99, column=1, pady=5)
    config.grid(row=98, column=1, pady=5)
    close.grid(row=99, column=3, pady=5)
    
    writefileB.grid(row=90, column=1, pady=20, sticky=(tk.E))
    brwslocW.grid(row=90, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBAWS.grid(row=86, column=1, pady=0, sticky=(tk.E))
    brwslocAWS.grid(row=86, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBPAS.grid(row=87, column=1, pady=0, sticky=(tk.E))
    brwslocPAS.grid(row=87, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBWIS.grid(row=88, column=1, pady=0, sticky=(tk.E))
    brwslocWIS.grid(row=88, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileBWSS.grid(row=89, column=1, pady=0, sticky=(tk.E))
    brwslocWSS.grid(row=89, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
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
def RunApplication(updatebox, window, Files):
    def callback():
        ConvertToXLSX(updatebox, window, Files)
        
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
    
    delimiterlbl = ttk.Label(ConfigWindow, text="Delimiter used in input file:", justify=tk.CENTER, padding=(12,12,12,0))
    headerstolbl = ttk.Label(ConfigWindow, text="Comma Separated List of Headers to Keep in Output File. \nThese should match EXACTLY those in the GradeX file. \nRestore Defaults by Deleting the custom-config.json file.", justify=tk.CENTER, padding=(12,24,12,0))
    forceheadlbl = ttk.Label(ConfigWindow, text="Include requested headers even if they can't be matched to one in the input file? \n(All row entries will be blank)", justify=tk.CENTER, padding=(12,24,12,0))
    delimiterent = ttk.Entry(ConfigWindow, width=5)
    delimiterent.insert(0, DELIMITER)
    headerstoent = ttk.Entry(ConfigWindow, width=75, text=headerssting)
    headerstoent.insert(0, headerssting)
    forceheadent = ttk.Checkbutton(ConfigWindow, variable=forcedhead)
    ok = ttk.Button(ConfigWindow, text="Save", command=lambda: WriteSettings(ConfigWindow, delimiterent.get(), headerstoent.get(), forcedhead.get()))
    close = ttk.Button(ConfigWindow, text="Close", command=ConfigWindow.destroy)
    
    
    delimiterlbl.grid(row=1, column=1, columnspan=2)
    delimiterent.grid(row=2, column=1, columnspan=2)
    headerstolbl.grid(row=3, column=1, columnspan=2)
    headerstoent.grid(row=4, column=1, columnspan=2, padx = 20)
    forceheadlbl.grid(row=5, column=1, columnspan=2)
    forceheadent.grid(row=6, column=1, columnspan=2)
    
    ok.grid(row=55, column=1, pady=20)
    close.grid(row=55, column=2, pady=20)
    
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
    if files[filename] == False or files[filename] == None: #check that files exist
        return -1
    if filename not in HEADERS and filename != "write": #check that headers are defined for file.
        return -2
    return 1

def _ProcessFiles(updatebox, window, files, name, workbooks):
    Data = {}
    filename = files[name].name
    global OUTSIDEKILL
    
    lineat = 0

    with open(filename, "r", encoding="utf-8-sig") as csvfile:
        csvrd = csv.reader(csvfile, delimiter=DELIMITER, quotechar=QUOTECHAR)
        for line in csvrd:
            #line = line.strip("\n").split(DELIMITER)
            if lineat == 0: #header business
                index = 0
                #figure out the location of all the headers.
                for header in line:
                    if header in HEADERS[name]:
                        HEADERS[name][header] = index
                    index += 1
                    
                for header in sorted(HEADERS[name]): #check if all headers matched
                    if HEADERS[name][header] < 0:
                        updateboxtext = "               WARNING!! Column \"" + header + "\" not found in file " + filename + "! " 
                        if header == "Region": updateboxtext+= "\n                 Region is required to build the files! Unable to process file!"
                        updatebox.insert(tk.END, updateboxtext)
                        updatebox.yview(tk.END)
                        window.update()
                        if FUZZYMATCHING: #If fuzzy matching enabled, try to match to something close
                            fuzzymatch = get_close_matches(header, line)
                            if len(fuzzymatch) > 0:
                                HEADERS[name][header] = line.index(fuzzymatch[0])
                                updateboxtext += "                 " + header + "column not found! Using closest match " + fuzzymatch[0]
                                updatebox.insert(tk.END, updateboxtext)
                                updatebox.yview(tk.END)
                                window.update()
                        if header == "Region" : return #important header, terminate
                        continue
                lineat += 1
                continue #read the next line of the file.
            #write to CSV file            
            
            workbookname = line[HEADERS[name]["Region"]]
            if workbookname == "" or workbookname == None:
                updateboxtext = "               Row #" + str(lineat) + " has no Region!"
                updatebox.insert(tk.END, updateboxtext)
                updatebox.yview(tk.END)
                continue
            
            if not workbookname in workbooks:
                wbn = files["write"] + "/" + workbookname + ".xlsx"
                wbf = xlsxwriter.Workbook(wbn)
                workbooks[workbookname] = XLSWorkbook(wbf, workbookname, name) 

            workbooks[workbookname].WriteLine(HEADERS[name], line, name)
                
                    
            
            '''
            for header in sorted(HEADERS[name]):
                if (FORCEHEADER and HEADERS[name][header] < 0): #check for unmatched headers.
                    if (row == 0):
                        worksheet.write(row, col, header, wbheaderformat)
                    col += 1
                
                if row == 0:
                    worksheet.write(row, col, line[HEADERS[name][header]],wbheaderformat)
                else:
                    worksheet.write(row, col, line[HEADERS[name][header]])
                col += 1
    
                
                
            '''    
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
        
    #updateboxtext = "      Finished Processing " + filename + ", total of " + str(lineat) + " rows" 
    #updatebox.insert(tk.END, updateboxtext)
    #updatebox.yview(tk.END)    
    #window.update()
    
def ConvertToXLSX(updatebox, window, files):
    global RUNNING
    global OUTSIDEKILL
    global FILETYPES
    RUNNING = True
    workbooks = {} #where we store the active workbooks as we write too them.
    
    try:            
        for name in FILETYPES:
            i = _CheckFiles(files, name)
            if i != 1: #error detected
                if i == -1:
                    updateboxtext = "Some files not selected! Please Check!"
                elif i == -2:
                    updateboxtext = "Configuration files are not correct! Please Check!"
                else:
                    updateboxtext = "General error while checking validity of files!"
                updatebox.insert(tk.END, updateboxtext)
                updatebox.yview(tk.END)    
                window.update()
                return #don't do anything if files invalid
                
            
        
        for name in FILETYPES: #iterate through the files
            if name == "write": continue
            filename = files[name].name
            files[name].close()
            updateboxtext = "      Reading File " + filename 
            updatebox.insert(tk.END, updateboxtext)
            updatebox.yview(tk.END)    
            window.update()
            _ProcessFiles(updatebox, window, files, name, workbooks) #load all files into memory
            
        
        
        updateboxtext = "File converted, output saved to " + files["write"]
        updatebox.insert(tk.END, updateboxtext)
        updatebox.yview(tk.END)    
        window.update()
    except:
        raise
    finally:
        #close all the workbooks (if any) open
        RUNNING = False
        if len(workbooks) > 0:
            for wbname in workbooks:
                workbooks[wbname].close()
            
            

def WriteSettings(ConfigWindow, delimiter, header, forcehead):
    global DELIMITER
    global HEADERS
    global FORCEHEADER
    DELIMITER = delimiter
    
    header = header.split(",")
    HEADERS = {}
    for item in header:
        HEADERS[item] = -1
    
    FORCEHEADER = forcehead
    
    with open(CONFIGFILE, "w") as configfile:
        settings = {"DELIMITER":DELIMITER, "HEADERS":HEADERS,"FORCEHEADER":FORCEHEADER}
        json.dump(settings, configfile, sort_keys=True, indent = 4)
    
    ConfigWindow.destroy()
    messagebox.showinfo(title="Configurations Saved", message='Configuration File saved as "custom-config.json"')
    
     
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
    main()
        
