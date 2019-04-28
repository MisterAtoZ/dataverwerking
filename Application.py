import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import Main

class Application(tk.Frame):

    def __init__(self, master=None):
        """
        initialisation of the gui
        """
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        """
        create the different widgets of the gui and put them in rows and columns
        """
        self.grid()
        self.pack(fill="both", expand=True)

        #tabs
        tabControl = ttk.Notebook(self)
        tabControl.pack(expand=1, fill='both')

        tabBifi = ttk.Frame(tabControl)
        tabControl.add(tabBifi, text='BIFI')

        tabPsc = ttk.Frame(tabControl)
        tabControl.add(tabPsc, text='PSC')

        tabSm = ttk.Frame(tabControl)
        tabControl.add(tabSm, text='SM')


        #variables
        self.hourVar = tk.IntVar()
        self.fileVar = tk.StringVar()
        self.comboVar = tk.StringVar()
        self.comboList = []     # nr + configContent
        self.configContent = [] # filepath
        self.samples = []       # frame + filepath
        self.ivCb = tk.IntVar()
        self.eqeCb = tk.IntVar()
        self.photoCb = tk.IntVar()

        self.filePsc = tk.StringVar()
        self.comboPsc = tk.StringVar()
        self.comboListPsc = ['']
        self.configPsc = []
        self.samplesPsc = []

        self.panelVar = tk.IntVar()
        self.fileSm = tk.StringVar()
        self.comboSm = tk.StringVar()
        self.comboListSm = [''] # nr + configContent
        self.configSm = []      # filepath

        #set variables if config file excists
        try:
            self.readFile()
        except IOError:
            self.comboList.append('') # voor Psc en Sm ook -----------------------------
            pass

        #labels
        self.hourLabel = tk.Label(tabBifi, text='Maximum  duration of stressing? [h]').grid(sticky='W', row=0, column=3,columnspan=2)
        self.configLabel = tk.Label(tabBifi, text='Select previous configuration').grid(sticky='W', row=1, column=3)
        self.fileLabel = tk.Label(tabBifi, text='Excel file').grid(sticky='W',row=2, column=3)
        self.errorLabel = tk.Label(tabBifi, text='', fg='red')
        self.errorLabel.grid(sticky='W', row=4, column=1,columnspan=3)

        self.configLabelPsc = tk.Label(tabPsc, text='Select previous configuration').grid(sticky='W', row=0, column=3)
        self.fileLabelPsc = tk.Label(tabPsc, text='Excel file').grid(sticky='W', row=1, column=3)
        self.errorLabelPsc = tk.Label(tabPsc, text='', fg='red')
        self.errorLabelPsc.grid(sticky='W', row=2, column=1, columnspan=3)

        self.panelLabel = tk.Label(tabSm, text='Number of panels').grid(sticky='W', row=0, column=3,columnspan=2)
        self.configLabelSm = tk.Label(tabSm, text='Select previous configuration').grid(sticky='W', row=1, column=3)
        self.fileLabelSm = tk.Label(tabSm, text='Excel file').grid(sticky='W', row=2, column=3)
        self.errorLabelSm = tk.Label(tabSm, text='', fg='red')
        self.errorLabelSm.grid(sticky='W', row=4, column=1, columnspan=3)

        #entries
        self.hourInput = tk.Entry(tabBifi, textvariable=self.hourVar)
        self.hourInput.grid(sticky='WE', row=0, column=0,columnspan=3,padx=5)
        self.fileInput = tk.Entry(tabBifi,textvariable=self.fileVar)
        self.fileInput.grid(row=2,column=0,columnspan=3,sticky='WE',padx=5)

        self.fileInputPsc = tk.Entry(tabPsc, textvariable=self.filePsc)
        self.fileInputPsc.grid(row=1, column=0, columnspan=3, sticky='WE', padx=5)

        self.panelInput = tk.Entry(tabSm, textvariable=self.panelVar)
        self.panelInput.grid(sticky='WE', row=0, column=0, columnspan=3, padx=5)
        self.fileInputSm = tk.Entry(tabSm, textvariable=self.fileSm)
        self.fileInputSm.grid(row=2, column=0, columnspan=3, sticky='WE', padx=5)

        #checkboxes
        self.iv = tk.Checkbutton(tabBifi, text="IV", variable=self.ivCb, onvalue=1, offvalue=0).grid(row=3, column=0, sticky='W')
        self.eqe = tk.Checkbutton(tabBifi, text="EQE", variable=self.eqeCb, onvalue=1, offvalue=0).grid(row=3, column=1, sticky='W')
        self.photo = tk.Checkbutton(tabBifi, text="Photo", variable=self.photoCb, onvalue=1, offvalue=0).grid(row=3, column=2,sticky='W')

        #buttons
        self.beginBtn = tk.Button(tabBifi, text='Begin', command=self.beginBifi).grid(sticky='W',row=4,column=0,padx=5)
        self.fileBtn = tk.Button(tabBifi, text='Pick file')
        self.fileBtn.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtn.grid(sticky='W', row=2, column=4)
        self.rmBtn = tk.Button(tabBifi, text='Remove config')
        self.rmBtn.bind('<ButtonRelease-1>', self.remove)
        self.rmBtn.grid(sticky='W', row=1, column=4)

        self.beginBtnPsc = tk.Button(tabPsc, text='Begin', command=self.beginPsc).grid(sticky='W', row=2, column=0, padx=5)
        self.fileBtnPsc = tk.Button(tabPsc, text='Pick file')
        self.fileBtnPsc.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtnPsc.grid(sticky='W', row=1, column=4)
        self.rmBtnPsc = tk.Button(tabPsc, text='Remove config')
        self.rmBtnPsc.bind('<ButtonRelease-1>', self.remove)
        self.rmBtnPsc.grid(sticky='W', row=0, column=4)

        self.beginBtnSm = tk.Button(tabSm, text='Begin', command=self.beginSm).grid(sticky='W', row=4, column=0, padx=5)
        self.fileBtnSm = tk.Button(tabSm, text='Pick file')
        self.fileBtnSm.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtnSm.grid(sticky='W', row=2, column=4)
        self.rmBtnSm = tk.Button(tabSm, text='Remove config')
        self.rmBtnSm.bind('<ButtonRelease-1>', self.remove)
        self.rmBtnSm.grid(sticky='W', row=1, column=4)

        #combobox
        self.combo = ttk.Combobox(tabBifi,textvariable=self.comboVar,values=self.comboList)
        self.combo.grid(row=1, column=0,columnspan=3, sticky='WE',padx=5)
        self.combo.current(len(self.comboList)-1)
        self.combo.bind('<<ComboboxSelected>>', self.select)

        self.comboP = ttk.Combobox(tabPsc,textvariable=self.comboPsc,values=self.comboListPsc)
        self.comboP.grid(row=0, column=0,columnspan=3, sticky='WE',padx=5)
        self.comboP.current(len(self.comboListPsc)-1)
        self.comboP.bind('<<ComboboxSelected>>', self.select)

        self.comboS = ttk.Combobox(tabSm, textvariable=self.comboSm, values=self.comboListSm)
        self.comboS.grid(row=1, column=0, columnspan=3, sticky='WE', padx=5)
        self.comboS.current(len(self.comboListSm) - 1)
        self.comboS.bind('<<ComboboxSelected>>', self.select)

        #configure column 2 to stretch with the window
        tabBifi.grid_columnconfigure(2, weight=1)
        tabPsc.grid_columnconfigure(2, weight=1)
        tabSm.grid_columnconfigure(2, weight=1)

    def getFrame(self, caller):
        if 'frame.' in str(caller):
            return 'Bifi'
        if 'frame2.' in str(caller):
            return 'Psc'
        if 'frame3.' in str(caller):
            return 'Sm'

    def remove(self, event):
        """
        removes a configuration form the config.txt file
        is called by rmBtn
        """
        caller = event.widget
        frame = self.getFrame(caller)

        if frame == 'Bifi':
            i = int(self.comboVar.get().split(' - ')[0]) - 1
            with open('config.txt','r') as file:
                filedata = file.read()
            filedata= filedata.replace('Bifi|' + self.configContent[i] + '\n', '')
            with open('config.txt', 'w') as file:
                file.write(filedata)
        if frame == 'Psc':
            i = int(self.comboPsc.get().split(' - ')[0]) - 1
            with open('config.txt', 'r') as file:
                filedata = file.read()
            filedata = filedata.replace('Psc|' + self.configPsc[i] + '\n', '')
            with open('config.txt', 'w') as file:
                file.write(filedata)

    def pickFile(self, event):
        """
        select a .xlsx file via the filedialog
        is called by fileBtn
        """
        caller = event.widget
        frame = self.getFrame(caller)

        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.filename != None:
            if frame == 'Bifi':
                self.fileVar.set(self.filename)
            if frame == 'Psc':
                self.filePsc.set(self.filename)
            if frame == 'Sm':
                self.fileSm.set(self.filename)

    def select(self, event):
        """
        select a configuration with the comboBox and set the configuration variables in the gui
        """
        caller = event.widget
        frame = self.getFrame(caller)
        if frame == 'Bifi':
            print(self.comboVar.get())
            path = self.comboVar.get().split(' - ')[1]
            self.setVars('Bifi', path)
        if frame == 'Psc':
            print(self.comboPsc.get())
            path = self.comboPsc.get().split(' - ')[1]
            self.setVars('Psc', path)
        if frame == 'Sm':
            print(self.comboSm.get())
            path = self.comboSm.get().split(' - ')[1]
            self.setVars('Sm', path)

    def readFile(self):
        """
        read the config.txt file and put the configurations in the comboBox and in the gui
        """
        with open('config.txt', 'r') as f:
            data = f.read()
            if data != '':
                self.sampels = data.split('\n')
                for i in range(0, len(self.sampels) - 1, 1):
                    frame = self.sampels[i].split('|')[0]
                    path = self.sampels[i].split('|')[1]
                    if frame == 'Bifi':
                        self.configContent.append(path)
                        self.setVars(frame, path)
                        self.comboVar.set(str(len(self.configContent)) + ' - ' + self.fileVar.get())
                        self.comboList.append(self.comboVar.get())
                    if frame == 'Psc':
                        self.configPsc.append(path)
                        self.setVars(frame, path)
                        self.comboPsc.set(str(len(self.configPsc)) + ' - ' + self.filePsc.get())
                        self.comboListPsc.append(self.comboPsc.get())
            else:
                self.comboList.append('')


    def setVars(self, frame, path):
        """
        set the variables from the configuration from the comboBox in the gui
        :param i: the number of the configuration
        """
        if frame == 'Bifi':
            self.fileVar.set(path)
        if frame == 'Psc':
            self.filePsc.set(path)
        if frame == 'Sm':
            self.filePsc.set(path)

    def beginBifi(self):
        """
        check if all variables are set correctly and starts the beginBifi function of the Main class
        is called by beginBtn
        """
        #wbName
        filePath=''
        filenameSplit = self.fileInput.get().split('/')
        for i in range(0,len(filenameSplit)-1,1):
            filePath = filePath + str(filenameSplit[i]) + '/'
        wbName = filenameSplit[len(filenameSplit)-1]
        # check if all variables are set correctly
        # yes: begin
        # no: show error message
        if self.hourInput.get() == "":
            self.errorLabel.config(text='Error : No hour')
        elif self.hourInput.get().isdigit():
            print('uur ' + str(self.hourInput.get()))
            # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
            if(self.fileInput.get().endswith('.xlsx')):
                self.errorLabel.config(text='Running...')
                self.errorLabel.update()
                if Main.Main.beginBifi(self, self.hourInput.get(), wbName, filePath, self.ivCb.get(), self.eqeCb.get(), self.photoCb.get()):
                    self.errorLabel.config(text='File saved')

                    configText = 'Bifi|' + self.fileInput.get()

                    # check if the file is already in config.txt and add it if not
                    inFile = False
                    for i in range(0, len(self.configContent), 1):
                        if configText == 'Bifi|' + self.configContent[i]:
                            inFile = True

                    if inFile == False:
                        with open('config.txt','a+') as f:
                            f.write(configText)
                            f.write('\n')
                            self.configContent.append(filePath + wbName)
                else:
                    self.errorLabel.config(text='Error: Something went wrong')
            else:
                self.errorLabel.config(text='Error : Not a .xlsx file')
        else:
            self.errorLabel.config(text='Error : Hour must be a number')

    def beginPsc(self):
        """
        check if all variables are set correctly and starts the beginPsc function of the Main class
        is called by beginBtnPsc
        """
        #wbName
        filePath=''
        filenameSplit = self.fileInputPsc.get().split('/')
        for i in range(0,len(filenameSplit)-1,1):
            filePath = filePath + str(filenameSplit[i]) + '/'
        wbName = filenameSplit[len(filenameSplit)-1]
        # check if all variables are set correctly
        # yes: begin
        # no: show error message

        # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
        if(self.fileInputPsc.get().endswith('.xlsx')):
            self.errorLabelPsc.config(text='Running...')
            self.errorLabelPsc.update()
            if Main.Main.beginPsc(self, wbName, filePath):
                self.errorLabelPsc.config(text='File saved')

                configText = 'Psc|' + self.fileInputPsc.get()

                # check if the file is already in config.txt and add it if not
                inFile = False
                for i in range(0, len(self.configPsc), 1):
                    if configText == 'Psc|' + self.configPsc[i]:
                        inFile = True

                if inFile == False:
                    with open('config.txt','a+') as f:
                        f.write(configText)
                        f.write('\n')
                        self.configPsc.append(filePath + wbName)
            else:
                self.errorLabelPsc.config(text='Error: Something went wrong')
        else:
            self.errorLabelPsc.config(text='Error : Not a .xlsx file')

    def beginSm(self):
        """
        check if all variables are set correctly and starts the beginSm function of the Main class
        is called by beginBtnSm
        """
        #wbName
        filePath=''
        filenameSplit = self.fileInputSm.get().split('/')
        for i in range(0,len(filenameSplit)-1,1):
            filePath = filePath + str(filenameSplit[i]) + '/'
        wbName = filenameSplit[len(filenameSplit)-1]
        # check if all variables are set correctly
        # yes: begin
        # no: show error message

        if self.panelInput.get() == "":
            self.errorLabelSm.config(text='Error : No panel')
        elif self.panelInput.get().isdigit():
            if int(self.panelInput.get()) > 0 and int(self.panelInput.get()) <= 6:
                # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
                if(self.fileInputSm.get().endswith('.xlsx')):
                    self.errorLabelSm.config(text='Running...')
                    self.errorLabelSm.update()
                    if Main.Main.beginSm(self, self.panelInput.get(), wbName, filePath):
                        self.errorLabelSm.config(text='File saved')

                        # configText = 'Sm|' + self.fileInputSm.get()
                        #
                        # # check if the file is already in config.txt and add it if not
                        # inFile = False
                        # for i in range(0, len(self.configSm), 1):
                        #     if configText == 'Sm|' + self.configSm[i]:
                        #         inFile = True
                        #
                        # if inFile == False:
                        #     with open('config.txt','a+') as f:
                        #         f.write(configText)
                        #         f.write('\n')
                        #         self.configSm.append(filePath + wbName)

                    else:
                        self.errorLabelSm.config(text='Error: Something went wrong')
                else:
                    self.errorLabelSm.config(text='Error : Not a .xlsx file')
            else:
                self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')
        else:
            self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')

root = tk.Tk()
root.title("PID data automation")
app = Application(master=root)
app.mainloop()