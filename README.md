# Instructions
When the application is executed, a graphical user interface (GUI) appears on screen. This GUI consists of 3 tabs: Bifi, Psc, and Switch matrix. This text explains, sorted by tab, what the program expects as inputs and what you can expect as outputs.

## LOANA
### User interface
The GUI consists of two entries, a combobox, three checkboxes, and three buttons.

In the first entry, the user can input a number. This number is the maximum hour that will be put in the Excel file. The standard value of this entry is 0.

The combobox is used to easily chose previous files, without the need to find them in the filedialog. The combobox receives the file paths from the config.txt file. When a configuration is chosen, it can be removed from the config.txt file, and thus from the combobox, with the use of the 'Remove config' button.

With the 'Pick file' button, the second entry can be filled in. When the user presses on the 'Pick file' button, a filedialog window pops up. When the right Excel file is opened, the file path is displayed in the entry. 

The tree checkboxes, IV, EQE, and Photo, are used to choose which information needs to be processed. When IV is checked, the program will put all the IV data in the Excel file. When EQE is checked, the EQE data will be put in the Excel file. When Photo is checked, the EL-photos are put together in one .jpg image. If something is not checked, it will not be processed.

The 'Begin' button is used to start the data processing. If the 'Maximum duration of stressing [h]' entry is not filled in, the error message 'Error : No hour' will be displayed on the bottom of the gui. If a letter or a symbol is entered in this entry, the error message 'Error : Hour must be a number' will be displayed. If the file is not a .xlsx, the error message 'Error : Not a .xlsx file' will be displayed. The last error message that can occur is 'Error: Something went wrong'. This message is displayed when the function called returns False. When the function called returns True, the message 'File saved' will be displayed, and the file path is saved in the config.txt file. When the program is running, the message 'Running...' is displayed.


### Excel file
The Excel file, which is chosen in the gui, needs to be filled in in a specific way. One row is used to label the first 5 columns: 'Date', 'Time (out)', 'Time (in)', 'Interval [h]', 'Acc. Hours [h]'. In the row below, the first 3 columns need to be filled in, the last two are done automatically. Column A is used for the dates of the measurement (format: dd/mm/yyyy). Column B gives the time when the stressing stops (format hh:mm). Column C gives the time when the stressing starts again (format hh:mm). These are the only things that need to be filled in. It is important that the columns are in the right order and starting from column A. The first four labels can have a different name, however, the last one needs to be 'Acc. Hours [h]'. The starting row is not important, therefore, it is possible to put extra information (samples, stress condition, equipment, etc.) above the data. The sheet name with all this information needs to be named 'General'.

This Excel file needs to be in the same directory as the directories of the measurements.

When the program is run, a new Excel file is present in the same directory. The name of this file starts with the last number of stressed hours that is present in the file, and then a hyphen and the original name of the Excel file. This new file contains all the available data. 

After the General sheet, four more sheets are present: %PID, FF, Voc, Isc. These sheets only contain a graph. Further, al the panels have a sheet with only their name and a sheet with 'EQE_' and their name. The first sheet contains the data found in 'data-exchange.xlsx' and all the light and dark IV measurements, sorted by hours of stressing. Two graphs of the light and dark IV measurements are also present. The EQE sheet contains the EQE data, sorted by hours of stressing. The normalized EQE is also calculated. The EQE and normalized EQE graphs are present.

### Measurements
The measurement files are sorted in different directories. In the same directory as the Excel file, the directories are sorted and named after the hours of stressing. In this directory, everything is sorted by name. In this directory are the 'IQE' and 'IV' directories. In the 'IQE' directory is the .eqe file with the EQE measurement data. The 'IV' directory contains the .lgt and .drk files with the light and dark IV measurement data, respectively. The name directories should be named the same in all hour directories. 

If the .lgt, .drk, or .eqe file is not present, the program ignores it and goes on to the next measurement.

The 'data-exchange.xlsx' file of the last measurement needs to be present, in the same directory as the other Excel file, or in the last hours folder.

### EL-images
If there are EL-photos, they need to be in the hours directory. The names of these images need to end with the correct name, corresponding to a name directory, plus '.jpg' or '.JPG'. These images are put together in one .jpg image with all the photos and titles. From top to bottom, the titles are the names of the panels. From left to right, the titles are the number of stressed hours. This generated image is located in the same directory as the Excel file.

## Psc
### User interface
The GUI consists of a combobox, an entry, and three buttons.

The combobox is used to easily chose previous files, without the need to find them in the filedialog. The combobox receives the file paths from the config.txt file. When a configuration is chosen, it can be removed from the config.txt file, and thus from the combobox, with the use of the 'Remove config' button.

With the 'Pick file' button, the second entry can be filled in. When the user presses on the 'Pick file' button, a filedialog window pops up. When the right Excel file is opened, the file path is displayed in the entry. 

The 'Begin' button is used to start the data processing. If the file is not a .xlsx, the error message 'Error : Not a .xlsx file' will be displayed. The last error message that can occur is 'Error: Something went wrong'. This message is displayed when the function called returns False. When the function called returns True, the message 'File saved' will be displayed, and the file path is saved in the config.txt file. When the program is running, the message 'Running...' is displayed.

### Excel file
The Excel file, which is chosen in the gui, can be a blank Excel file. When making a new Excel file, the first sheet can be removed or it can be used to enter some information. The program will ignore this sheet and make new sheets for every sample. These sheets contain the IV data of the sample sorted by stressing time. A graph of these IV data is also made.

This Excel file needs to be in the same directory as the files of the measurements.

When the program is run, a new Excel file is present in the same directory. The name of this file starts with the last number of stressed hours that is present in the file, and then a hyphen and the original name of the Excel file. This new file contains all the available data. 

### Measurements
The measurement files are all in the same directory, in the same directory as the Excel file. The .dat files contain the IV data, the .ini files contain information about the configuration. The names of these .dat files need to have the format: sample name - amount of stressed minutes - sample number - measurement number. For example: Sample-30min-1-1.


## Switch matrix
### User interface
The GUI consists of a combobox, an entry, and three buttons.

The combobox is used to easily chose previous files, without the need to find them in the filedialog. The combobox receives the file paths from the config.txt file. When a configuration is chosen, it can be removed from the config.txt file, and thus from the combobox, with the use of the 'Remove config' button.

With the 'Pick file' button, the second entry can be filled in. When the user presses on the 'Pick file' button, a filedialog window pops up. When the right Excel file is opened, the file path is displayed in the entry. 

The 'Begin' button is used to start the data processing. If the file is not a .xlsx, the error message 'Error : Not a .xlsx file' will be displayed. The last error message that can occur is 'Error: Something went wrong'. This message is displayed when the function called returns False. When the function called returns True, the message 'File saved' will be displayed, and the file path is saved in the config.txt file. When the program is running, the message 'Running...' is displayed.

### Excel file
The Excel file, which is chosen in the gui, can be a blank Excel file. When making a new Excel file, the first sheet can be removed or it can be used to enter some information. The program will ignore this sheet and make new sheets for every module. These sheets contain the dark IV data of the module sorted by stressing time. A graph of these IV data is also made.

This Excel file needs to be in the same directory as the directories of the modules.

When the program is run, a new Excel file is present in the same directory. The name of this file starts with the last number of stressed hours that is present in the file, and then a hyphen and the original name of the Excel file. This new file contains all the available data. 

### Measurements
The measurement files are in the module directories. These .csv files contain the dark IV measurements of the module at a certain time. The name of the file gives the number of hours of stressing.