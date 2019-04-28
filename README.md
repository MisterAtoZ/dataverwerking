# Instructions
When the application is executed, a graphical user interface (GUI) appears on screen. This GUI consists of 3 tabs: Bifi, Psc, and Sm. This text explains, ordend by tab, what the program expects as inputs and what you can expect as outputs.

## Bifi
### User interface
The GUI consists of two entries, combobox, and three buttons.

In the first entry, the user can input a number. This number is the maximum hour that will be put in the Excel file. The standard value of this entry is 0.

With the 'Pick file' button, the second entry can be filled in. When the user presses on the 'pick file' button, a filedialog window pops up. When the right Excel file is opened, the file path is displayed in the entry. 

The 'Begin' button is used to start the data processing. If the 'Maximum duration of stressing [h]' entry is not filled in, the error message 'Error : No hour' will be displayed on the bottom of the gui. If a letter or a symbol is entered in this entry, the error message 'Error : Hour must be a number' will be displayed. If the file is not a .xlsx, the error message 'Error : Not a .xlsx file' will be displayed. The last error message that can occur is 'Error: Something went wrong'. This message is displayed when the function called returns False. When the function called returns True, the message 'File saved' will be displayed, and the file path is saved in the config.txt file.

The combobox is used to easily chose previous file paths, without the need to find them in the filedialog. The combobox receives the file paths from the config.txt file. When a configuration is chosen, it can be removed from the config.txt file, and thus from the combobox, with the use of the 'Remove config' button.

### Excel file
The Excel file, which is chosen in the gui, needs to be filled in in a specific way. One row is used to label the first 5 columns: 'Date', 'Time (out)', 'Time (in)', 'Interval [h]', 'Acc. Hours [h]'. In the row below, the first 3 columns need to be filled in, the last two are done automatically. Column A is used for the dates of the measurement (format: dd/mm/yyyy). Column B gives the time when the stressing stops (format hh:mm). Column C gives the time when the stressing starts again (format hh:mm). These are the only things that need to be filled in. It is important that the columns are in the right order and starting from column A. The first four labels can have a different name, however, the last one needs to be 'Acc. Hours [h]'. The starting row is not important, therefore, it is possible to put extra information (samples, stress condition, equipment, etc.) above the data. The sheet name with all this information needs to be named 'General'.

This Excel file needs to be in the same directory as the directories of the measurements.

When the program is run, a new Excel file is present in the same directory. This name of this file starts with the last number of stressed hours that is present in the file, and then an underscore and the original name of the Excel file. This new file contains all the available data. 

After the General tab, four more tabs are present: %PID, FF, Voc, Isc. These tabs only contain a graph. Further, al the panels have a tab with only their name and a tab with 'EQE_' and their name. The first tab contains the data found in 'data-exchange.xlsx' and all the light and dark IV measurements, sorted by hours of stressing. Two graphs of the light and dark IV measurements are also present. The EQE tab contains the EQE data, sorted by hours of stressing. The normalized EQE is also calculated. The EQE and normalized EQE graphs are present.

### Measurements
The measurement files are sorted in different directories. In the same directory as the Excel file, the directories are sorted and named after the hours of stressing. In this directory, everything is sorted by name. In this directory are the 'IQE' and 'IV' directories. In the 'IQE' directory is the .eqe file with the EQE measurement data. The 'IV' directory contains the .lgt and .drk files with the light and dark IV measurement data, respectively. The name directories should be named the same in all hour directories. 

If the .lgt, .drk, or .eqe file is not present, the program ignores it and goes on to the next measurement.

The 'data-exchange.xlsx' file of the last measurement needs to be present, in the same directory as the other Excel file, or in the last hours folder.

### EL-images
If there are EL-images, they need to be in the hours directory. The names of these images need to end with the correct name, corresponding to a name directory, plus '.jpg' or '.JPG'. These images are put together in one .jpg image with all the photos and titles. From top to bottom, the titles are the names of the panels. From left to right, the titles are the number of stressed hours. 

## Psc
### User interface

### Excel file

### Measurements



## Sm
### User interface

### Excel file

### Measurements
