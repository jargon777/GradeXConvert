import os
import tkinter


DELIMITER = ","
HEADERS = ["Loc Id", 
           "Risk - Total", 
           "Region", 
           "RWY",
           "Mile",
           "Subdivision Name",
           "Spur Mile",
           "Spur Name",
           "Date Inspected",
           "Inspected By",
           "Protection Type"]

def main():
    #MainWindow = tkinter.Tk()
    #MainWindow.mainloop()
    writedir = "converted"
    
    for filename in os.listdir("."):
        if filename.endswith("csv"):
            if not os.path.isdir(writedir):
                os.makedirs(writedir)
            
            print("Reading " + filename)
            with open(filename, "r", encoding="utf-8-sig") as csvfile:
                lineat = 0
                for line in csvfile:
                    line = line.split(DELIMITER)
                    if lineat == 0: #header
                        headers = line
                        indexes = []
                        print(headers)
                        for header in headers:
                            pass
                        
                        
                        
                    lineat += 1

if __name__ == "__main__":
    try:
        main()
        
    except:
        print("Unhandled Error.")
        raise