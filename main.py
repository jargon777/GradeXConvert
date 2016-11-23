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
RUNNING = False
OUTSIDEKILL = False

MainWindow = tk.Tk()
MainWindow.title("GradeXConvertToXLSX")
MainWindow.protocol('WM_DELETE_WINDOW', lambda: CloseProgram(MainWindow, None))


def main():
    Files = {"read": False, "write": False}
    
    settings = ReadSettings() #check for settings file and load.
    #widget def
    config = ttk.Button(MainWindow, text="Configure", command=ShowConfig)
    textlbl = ttk.Label(MainWindow, text='Application Messages'
                            ,width=75, wraplength=550, justify=tk.LEFT, padding=(12,12,12,12))
    messagelist = tk.Listbox(MainWindow, height=3, width=70)
    
    ok = ttk.Button(MainWindow, text="Run", command=lambda: RunApplication(messagelist, MainWindow, Files))
    close = ttk.Button(MainWindow, text="Cancel", command=lambda: CloseProgram(MainWindow, messagelist))
    
    brwslocR = ttk.Entry(MainWindow)
    brwslocR.insert(0, "Please Load a GradeX Output File")
    brwslocR.configure(state='disabled')
    brwslocW = ttk.Entry(MainWindow)
    brwslocW.insert(0, "Please Choose a File to Save to")
    brwslocW.configure(state='disabled')
    readfileB = ttk.Button(MainWindow, text="Convert...", command=lambda: askFile("read", Files, brwslocR, MainWindow))
    writefileB = ttk.Button(MainWindow, text="Save to...", command=lambda: askFile("write", Files, brwslocW, MainWindow))
    #widget layout
    #textlbl.grid(row=0, column=1, columnspan=3)
    messagelist.grid(row=2, column=1, columnspan=3, sticky=(tk.N, tk.S, tk.E, tk.W), pady=20)
    messagelist.insert(tk.END, "GradeX Output Converter")
    messagelist.insert(tk.END, "  Select a File to Convert...")
    
    ok.grid(row=99, column=1, pady=20)
    config.grid(row=99, column=2, pady=20)
    close.grid(row=99, column=3, pady=20)
    
    writefileB.grid(row=98, column=1, pady=0, sticky=(tk.E))
    brwslocW.grid(row=98, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
    readfileB.grid(row=97, column=1, pady=0, sticky=(tk.E))
    brwslocR.grid(row=97, column=2, pady=0, columnspan = 2, sticky=(tk.E, tk.W))
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

def ConvertToXLSX(updatebox, window, files):
    global RUNNING
    global OUTSIDEKILL
    RUNNING = True
    
    if files["read"] == False or files["read"] == None or files["write"] == None or files["write"] == False:
        updateboxtext = "Some files not selected! Please Check!"
        updatebox.insert(tk.END, updateboxtext)
        updatebox.yview(tk.END)    
        window.update()
        RUNNING = False
        return #don't do anything if files invalid
    
    filename = files["read"].name
    files["read"].close()
    updateboxtext = "      Reading File " + filename 
    updatebox.insert(tk.END, updateboxtext)
    updatebox.yview(tk.END)    
    window.update()
    
    workbookname = files["write"].name
    files["write"].close() #close the file, we just want the location.
    workbook = xlsxwriter.Workbook(workbookname)
    worksheet = workbook.add_worksheet()
    wbheaderformat = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
            
    lineat = 0
    with open(filename, "r", encoding="utf-8-sig") as csvfile:
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
                if HEADERS[header] < 0:
                    if not lineat: #notify of bad headers at start.
                        updateboxtext = "               WARNING!! Column \"" + header + "\" not found!" 
                        updatebox.insert(tk.END, updateboxtext)
                        updatebox.yview(tk.END)    
                        window.update()
                    continue
                
                if row == 0:
                    worksheet.write(row, col, line[HEADERS[header]],wbheaderformat)
                else:
                    worksheet.write(row, col, line[HEADERS[header]])
                col += 1
    
                
                
                
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
    
    workbook.close()
    
    updateboxtext = "File converted, output saved as " + workbookname
    updatebox.insert(tk.END, updateboxtext)
    updatebox.yview(tk.END)    
    window.update()
            
    RUNNING = False

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
        #Calls a dialog box that asks the user to navigate to a folder to save localconfig.
        if key == "read":
            file = tkf.askopenfile("r", defaultextension=".csv", filetypes =(("GradeX Output", "*.csv"),("All Files","*.*")))
        if key == "write":
            file = tkf.asksaveasfile("w", defaultextension=".xlsx", filetypes =(("Microsoft Excel Table", "*.xlsx"),("All Files","*.*")))
        if file != False and file != None:
            fileDict[key] = file
            label.configure(state='normal')
            label.delete(0, tk.END)
            label.insert(0, file.name)
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
        
