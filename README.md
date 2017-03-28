# GradeXConvert
This program is written specifically to convert output from GradeX (http://gradex.ca/). This program works by taking output provided by this service (csv files) and strips unneeded columns, saving the result as a Microsoft Excel xlsx file.

This is a general purpose batch CSV to Excel converter that has the ability to remove unwanted columns.

## Using the Program
At the main screen, the top shows the messages the program is sending. There are 5 buttons to load files, a configure button, a run button and a close button. Load an AWS, PASSIVE, WIS, WSS CSV file using these buttons. Select a location to save using the “Save to…” button. When ready press “Run”. While the programming is running, pressing “Close” suspends the active processes (will not close the application).

## Configuration
Configuration must be done manually through a JSON file. A number of options can be configured. In the program, the “configure” button can be used to navigate to a window that can export the default configuration file and load a file after it’s been changed. The following options can be configured:

### ALLOWPARTIALFILE
Boolean. Whether or not the program will run if some of the files are missing. Default is false, meaning AWS, PASSIVE, WIS and WSS CSV inputs must be provided.

### DELIMTER 
The delimiter used to separate values in the input CSV file.

### FILENAMEPREPEND 
This text is prepended to the filenames of the workbooks written, e.g. if it is set to 'test' then the AWS workbook will be 'testAWS.xlsx'. These names are not santized, and invalid characters will throw a file name error on execution.

### FORCEHEADER 
Boolean. If a suitable match for the column can't be found, then should the column be kept?

### FUZZYLIMIT 
Integer between 0 and 1, configures the limit of the fuzzy matching. 1 means more strict matching, 0 less strict.

### FUZZYMATCHING 
Boolean. If set to true, the program will attempt to find a close match for the requested column. If set to false, exact matches are needed to column names. Fuzzy matching issues warnings to the screen if used.

### HEADERS 
The headers included in the input files. The overall format is a dictionary datatype, declared as HEADERS['TAB']['NAME IN INPUT'] = {'pos':1, 'remap': NAME IN OUTPUT}. In laymans terms, this setting has three levels. In the first level, each of the tags is used to generate separate tabs in the files. This cannot be changed in the setting file without throwing errors. The second level has the names of the columns from the source files that will be included in the written file. Be sure the column exists if you change these values. The third level has two keys, 'pos' and 'remap'. 'pos' is used when converting, and must take the form 'pos': -1. 'remap' controls what is written to the ouput file. Existing headers can be renamed. The order of entries determines the order in the output.

### PROVINCEPOSTALCONV 
Boolean. Whether or not the program will try to convert the province names from non-standard abbreviations to postal abbreviations

### PROVINCEPREMAP 
The names that are converted, in format 'Abbreviation : Postal Abbreviation'. Values matching 'Abbreviuation' will be converted to 'Postal Abbreviation'

### QUOTECHAR 
The quoting used to group together values that the delimiter in them in the input CSV file.

### SUMMATIONCOL 
Name of the column to use when generating the RWY sums

### SUMMATIONDO 
Boolean. Whether or not the summation statistics are written to the excel file

### ALLOWPARTIALFILE
Boolean. Whether or not the program will run if some of the files are missing. Default is false, meaning AWS, PASSIVE, WIS and WSS CSV inputs must be provided.

### ALLOWPARTIALFILE
Boolean. Whether or not the program will run if some of the files are missing. Default is false, meaning AWS, PASSIVE, WIS and WSS CSV inputs must be provided.

###DELIMTER 
The delimiter used to separate values in the input CSV file.

###FILENAMEPREPEND 
This text is prepended to the filenames of the workbooks written, e.g. if it is set to 'test' then the AWS workbook will be 'testAWS.xlsx'. These names are not santized, and invalid characters will throw a file name error on execution.

###FORCEHEADER 
Boolean. If a suitable match for the column can't be found, then should the column be kept?

###FUZZYLIMIT 
Integer between 0 and 1, configures the limit of the fuzzy matching. 1 means more strict matching, 0 less strict.

###FUZZYMATCHING 
Boolean. If set to true, the program will attempt to find a close match for the requested column. If set to false, exact matches are needed to column names. Fuzzy matching issues warnings to the screen if used.

###HEADERS 
The headers included in the input files. This setting has two levels. In the first level, each of the tags is used to generate separate tabs in the files. This cannot be changed in the setting file without throwing errors. The second level has the names of the columns from the source files that will be included in the written file. Be sure the column exists if you change these values. Column labels that you want included must take the form 'NAME': -1. Existing headers can be renamed.

###PROVINCEPOSTALCONV 

Boolean. Whether or not the program will try to convert the province names from non-standard abbreviations to postal abbreviations

###PROVINCEPREMAP 
The names that are converted, in format 'Abbreviation : Postal Abbreviation'. Values matching 'Abbreviuation' will be converted to 'Postal Abbreviation'

###QUOTECHAR 
The quoting used to group together values that the delimiter in them in the input CSV file.

###SUMMATIONCOL 
Name of the column to use when generating the RWY sums

###SUMMATIONDO 
Boolean. Whether or not the summation statistics are written to the excel file

###SPLITFILEON 
String, the column that is used to separate the inputs into different files. Default is to create region-based files.

###DONOTSPLITFILE
Files that will not be split according to SPLIFILEON

