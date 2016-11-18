# GradeXConvert
This program is written specifically to convert output from GradeX (http://gradex.ca/). This program works by taking output provided by this service (csv files) and strips unneeded columns, saving the result as a Microsoft Excel xlsx file.

This is a general purpose batch CSV to Excel converter that has the ability to remove unwanted columns.

## Using the Program
This program will only convert files that are in the same directory as the excecutable. Place files you wish to convert in the program's directory, and then open the application. At the main interface, press the "Run" button to begin the conversion. The messagebox is periodically updated as the program runs. When finished, a message box will be displayed summarizing what was done.

## Configuration
Configuration of the output is possible using the "Configure" button on the main UI. After pressing this button, a few configurable options are displayed. When you have made the modifications required, press the "save" button. Configuration is updated immediately in the program and custom configuration file, "custom-config.json". To restore default configurations, delete this file. A summary of the configurations avaialble is shown below.

### Configuring the Delimiter
This is the delimiter used in the input data files to separate data entries. This should not need to be changed, and is usually a comma. 

### Configuring the List of Headers
This is a comma seperated list of all the column names in the source file that will be saved. A default list is included with the program. You can add names to the list. These names are CASE SENSITIVE and must be entered exactly as they are in the source data (including spaces).

### Configuring the Behaviour of Unmatched Columns
By default, this converter will create a column for all requested data labels, even if they are not in the data. This behaviour can be altered by unticking the checkbox in the configuration panel. The program will then only include columns that have been successfully matched.
