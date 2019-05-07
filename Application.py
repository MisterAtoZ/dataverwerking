import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
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

        #styles
        bgColor = '#343434' #29648A'
        fgColor = '#FFFFFF' #DFDFDF'
        red = '#f91919'
    #     gui_style = ttk.Style()
    #     #gui_style.configure('My.TButton', background='#2E9CCA')
    #     gui_style.map("My.TButton",
    # #foreground=[('pressed', 'red'), ('active', bgColor)],
    # background=[('pressed', '!disabled', bgColor), ('active', bgColor), ('!active', bgColor)]
    # )
    #     gui_style.configure('My.TFrame', background=bgColor)
    #     gui_style.configure('My.TNotebook', background=fgColor)
    #     gui_style.configure('My.TLabel', background=bgColor, foreground=fgColor)
    #     gui_style.configure('My.TCheckbutton', background=bgColor, foreground=fgColor)
    #     gui_style.configure('My.TLabelframe.Label', background=bgColor, foreground=fgColor)
    #     gui_style.configure('My.TRadiobutton', background=bgColor, foreground=fgColor)
    #     gui_style.configure('My.TSeparator', background='#000000')
        # #style
        # style = ttk.Style()
        # style.theme_create('appstyle', parent='alt',
        #                    settings={
        #                        'TLabelframe': {
        #                            'configure': {
        #                                'background': bgColor,
        #                                'bd': 'red'
        #                            }
        #                        },
        #                        'TLabelframe.Label': {
        #                            'configure': {
        #                                'background': bgColor,
        #                                'foreground': fgColor,
        #                                'bd': 'red'
        #                            }
        #                        }
        #                    }
        #                    )
        # style.theme_use('appstyle')

        #tabs
        tabControl = ttk.Notebook(self, style='My.TNotebook')
        tabControl.pack(expand=1, fill='both')

        tabBifi = ttk.Frame(tabControl, style='My.TFrame')
        tabControl.add(tabBifi, text='Ex-situ')

        # tabPsc = ttk.Frame(tabControl, style='My.TFrame')
        # tabControl.add(tabPsc, text='Thin film')

        tabSm = ttk.Frame(tabControl, style='My.TFrame')
        tabControl.add(tabSm, text='In-situ')


        #variables
        self.hourRb = tk.IntVar()
        self.exSituRb = tk.IntVar()
        self.hourVar = tk.IntVar()
        self.fileVar = tk.StringVar()
        self.comboVar = tk.StringVar()
        self.comboList = []     # nr + file
        self.configContent = [] # nr + filepath
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

        #Ex-situ
            # Machine choise
        ttk.Separator(tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=0, columnspan=5, sticky='WE')
        self.machineLabel = ttk.Label(tabBifi, text='Select machine', style='My.TLabel').grid(sticky='W', row=0, column=0, columnspan=2, padx=5)
        self.exSituRbLoana = ttk.Radiobutton(tabBifi, text='LOANA', variable=self.exSituRb, value=1,style='My.TRadiobutton',command = lambda:self.selectMachine())\
            .grid(row=1,sticky='W', padx=5)
        self.exSituRbThin = ttk.Radiobutton(tabBifi, text='Thin film', variable=self.exSituRb, value=2,style='My.TRadiobutton',command=lambda: self.selectMachine())\
            .grid(row=1, column=1,sticky='W')
        self.exSituRb.set(1)
        ttk.Separator(tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=2, columnspan=5, sticky='WE')
            # Hour choise
        self.hourLabel = ttk.Label(tabBifi, text='Maximum  duration of stressing? [h]', style='My.TLabel').grid(sticky='W', row=2, column=0, columnspan=3, padx=5)
        self.hourRbAll = ttk.Radiobutton(tabBifi, text='All', variable=self.hourRb, value=1,style='My.TRadiobutton').grid(row=3, sticky='W', padx=5)
        self.hourRbEntry = ttk.Radiobutton(tabBifi, text='Hours:', variable=self.hourRb, value=2,style='My.TRadiobutton').grid(row=3, column=1,sticky='W')
        self.hourRb.set(1)
        self.hourInput = ttk.Entry(tabBifi, textvariable=self.hourVar)
        self.hourInput.grid(sticky='WE', row=3, column=2, pady=5,padx=(0,5))
        ttk.Separator(tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=4, columnspan=5, sticky='WE')
            # Excel file
        self.excelLabel = ttk.Label(tabBifi, text='Select Excel file', style='My.TLabel').grid(sticky='W',row=4,column=0,columnspan=2, padx=5)
        #self.configLabel = ttk.Label(tabBifi, text='Select previous configuration', style='My.TLabel').grid(sticky='W',row=1,column=3)
        self.combo = ttk.Combobox(tabBifi, textvariable=self.comboVar, values=self.comboList)
        self.combo.grid(row=5, column=0, columnspan=3, sticky='WE', padx=5)
        self.combo.current(len(self.comboList) - 1)
        self.combo.bind('<<ComboboxSelected>>', self.select)
        self.rmBtn = ttk.Button(tabBifi, text='Remove config', style='My.TButton')
        self.rmBtn.bind('<ButtonRelease-1>', self.remove)
        self.rmBtn.grid(sticky='WE', row=5, column=4, padx=5)
        self.fileInput = ttk.Entry(tabBifi, textvariable=self.fileVar)
        self.fileInput.grid(row=6, column=0, columnspan=3, sticky='WE', padx=5, pady=5)
        self.fileBtn = ttk.Button(tabBifi, text='Pick file', style='My.TButton')
        self.fileBtn.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtn.grid(sticky='WE', row=6, column=4, padx=5)
            # Data choise
        ttk.Separator(tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=7, columnspan=5, sticky='WE')
        self.dataLabel = ttk.Label(tabBifi, text='Select data types', style='My.TLabel').grid(sticky='W', row=7, column=0,columnspan=2, padx=5)
        self.iv = ttk.Checkbutton(tabBifi, text="IV", variable=self.ivCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.iv.grid(row=8, column=0, sticky='W', padx=5)
        self.eqe = ttk.Checkbutton(tabBifi, text="EQE", variable=self.eqeCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.eqe.grid(row=8, column=1, sticky='W')
        self.photo = ttk.Checkbutton(tabBifi, text="Photo", variable=self.photoCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.photo.grid(row=8, column=2, sticky='W')
            # Start
        self.beginBtn = ttk.Button(tabBifi, text='Begin', command=self.beginBifi, style='My.TButton')
        self.beginBtn.grid(sticky='W',row=9,column=0,padx=5,pady=5)
        self.beginBtn.config(width=9)
        self.errorLabel = ttk.Label(tabBifi, text='', foreground='red', font=('verdana', 10, ''))#, background=bgColor, foreground=red, font=('verdana', 10, ''))
        self.errorLabel.grid(sticky='W', row=9, column=1, columnspan=3)




        #labels


        #self.fileLabel = ttk.Label(tabBifi, text='Excel file', style='My.TLabel').grid(sticky='W',row=2, column=3)


        # self.configLabelPsc = ttk.Label(tabPsc, text='Select previous configuration', style='My.TLabel').grid(sticky='W', row=0, column=3)
        # self.fileLabelPsc = ttk.Label(tabPsc, text='Excel file', style='My.TLabel').grid(sticky='W', row=1, column=3)
        # self.errorLabelPsc = ttk.Label(tabPsc, text='', background=bgColor, foreground=red)
        # self.errorLabelPsc.grid(sticky='W', row=2, column=1, columnspan=3)

        # self.panelLabel = tk.Label(tabSm, text='Number of panels').grid(sticky='W', row=0, column=3,columnspan=2)
        self.configLabelSm = ttk.Label(tabSm, text='Select previous configuration', style='My.TLabel').grid(sticky='W', row=1, column=3)
        self.fileLabelSm = ttk.Label(tabSm, text='Excel file', style='My.TLabel').grid(sticky='W', row=2, column=3)
        self.errorLabelSm = ttk.Label(tabSm, text='', background=bgColor, foreground=red)
        self.errorLabelSm.grid(sticky='W', row=4, column=1, columnspan=3)


        #entries



        # self.fileInputPsc = ttk.Entry(tabPsc, textvariable=self.filePsc)
        # self.fileInputPsc.grid(row=1, column=0, columnspan=3, sticky='WE', padx=5)

        # self.panelInput = tk.Entry(tabSm, textvariable=self.panelVar)
        # self.panelInput.grid(sticky='WE', row=0, column=0, columnspan=3, padx=5)
        self.fileInputSm = ttk.Entry(tabSm, textvariable=self.fileSm)
        self.fileInputSm.grid(row=2, column=0, columnspan=3, sticky='WE', padx=5)



        #buttons





        # self.beginBtnPsc = ttk.Button(tabPsc, text='Begin', command=self.beginPsc, style='My.TButton').grid(sticky='W', row=2, column=0, padx=5, pady=5)
        # self.fileBtnPsc = ttk.Button(tabPsc, text='Pick file', style='My.TButton')
        # self.fileBtnPsc.bind('<ButtonRelease-1>', self.pickFile)
        # self.fileBtnPsc.grid(sticky='WE', row=1, column=4, padx=5)
        # self.rmBtnPsc = ttk.Button(tabPsc, text='Remove config', style='My.TButton')
        # self.rmBtnPsc.bind('<ButtonRelease-1>', self.remove)
        # self.rmBtnPsc.grid(sticky='W', row=0, column=4, padx=5, pady=(5,2))

        self.beginBtnSm = ttk.Button(tabSm, text='Begin', command=self.beginSm, style='My.TButton').grid(sticky='W', row=4, column=0, padx=5, pady=5)
        self.fileBtnSm = ttk.Button(tabSm, text='Pick file', style='My.TButton')
        self.fileBtnSm.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtnSm.grid(sticky='WE', row=2, column=4, padx=5)
        self.rmBtnSm = ttk.Button(tabSm, text='Remove config', style='My.TButton')
        self.rmBtnSm.bind('<ButtonRelease-1>', self.remove)
        self.rmBtnSm.grid(sticky='W', row=1, column=4, padx=5, pady=(5,2))

        #combobox


        # self.comboP = ttk.Combobox(tabPsc,textvariable=self.comboPsc,values=self.comboListPsc)
        # self.comboP.grid(row=0, column=0,columnspan=3, sticky='WE',padx=5)
        # self.comboP.current(len(self.comboListPsc)-1)
        # self.comboP.bind('<<ComboboxSelected>>', self.select)

        self.comboS = ttk.Combobox(tabSm, textvariable=self.comboSm, values=self.comboListSm)
        self.comboS.grid(row=1, column=0, columnspan=3, sticky='WE', padx=5)
        self.comboS.current(len(self.comboListSm) - 1)
        self.comboS.bind('<<ComboboxSelected>>', self.select)

        #configure column 2 to stretch with the window
        tabBifi.grid_columnconfigure(2, weight=1)
        # tabPsc.grid_columnconfigure(2, weight=1)
        tabSm.grid_columnconfigure(2, weight=1)

        # #Setting it up
        # img = ImageTk.PhotoImage(Image.open("imec.png"))
        # # Displaying it
        # imglabel = ttk.Label(tabBifi, image=img).grid(row=5)

        img = Image.open("imec.png")
        basewidth = 60 #70 voor zelfde breedte
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
        img = img.resize((basewidth, hsize), Image.ANTIALIAS)
        img = ImageTk.PhotoImage(img)
        label1 = Label(tabBifi, image=img)
        label1.image = img
        label1.grid(row=9, column=4, padx=5, pady=(5,5), sticky='E')#, columnspan=2) #pady=(0,5) voor breedte
        #label1.place(x=20, y=20)


    def selectMachine(self):
        if self.exSituRb.get() == 1:
            self.iv.config(state=NORMAL)
            self.eqe.config(state=NORMAL)
            self.photo.config(state=NORMAL)
        else:
            self.iv.config(state=DISABLED)
            self.eqe.config(state=DISABLED)
            self.photo.config(state=DISABLED)

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
            #filepath = self.comboList[i].split('|')[1]
            filepath = self.configContent[i].split('-')[1]
            print(filepath)
            with open('config.txt','r') as file:
                filedata = file.read()
            filedata= filedata.replace('Bifi|' + filepath + '\n', '')
            with open('config.txt', 'w') as file:
                file.write(filedata)
        # if frame == 'Psc':
        #     i = int(self.comboPsc.get().split(' - ')[0]) - 1
        #     with open('config.txt', 'r') as file:
        #         filedata = file.read()
        #     filedata = filedata.replace('Psc|' + self.configPsc[i] + '\n', '')
        #     with open('config.txt', 'w') as file:
        #         file.write(filedata)

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
            #print(self.comboList[(int(self.comboVar.get().split(' - ')[0]) - 1)][0]) #----------------------------------
            #print(self.comboList[(self.comboVar.get().split(' - ')[0]) - 1].split('|')[1])
            nr = int(self.comboVar.get().split(' - ')[0])-1
            path = self.configContent[nr]
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
                        nr = len(self.configContent) + 1
                        self.configContent.append(str(nr) + '-' + path)
                        self.setVars(frame, path)
                        filenames = self.fileVar.get().split('/')
                        filename = filenames[len(filenames)-1]
                        self.comboVar.set(str(nr) + ' - ' + filename)
                        self.comboList.append(str(self.comboVar.get()))
                        # self.comboList.append(self.comboVar.get())
                    # if frame == 'Psc':
                    #     self.configPsc.append(path)
                    #     self.setVars(frame, path)
                    #     self.comboPsc.set(str(len(self.configPsc)) + ' - ' + self.filePsc.get())
                    #     self.comboListPsc.append(self.comboPsc.get())
            else:
                self.comboList.append('')
            print(self.comboVar.get())


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
                if(self.hourRb.get() == 1):
                    hours = -1
                else:
                    hours = self.hourInput.get()
                if Main.Main.beginBifi(self, hours, wbName, filePath, self.ivCb.get(), self.eqeCb.get(), self.photoCb.get()):
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

        # if self.panelInput.get() == "":
        #     self.errorLabelSm.config(text='Error : No panel')
        # elif self.panelInput.get().isdigit():
        #     if int(self.panelInput.get()) > 0 and int(self.panelInput.get()) <= 6:
        #         # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
        if(self.fileInputSm.get().endswith('.xlsx')):
            self.errorLabelSm.config(text='Running...')
            self.errorLabelSm.update()
            if Main.Main.beginSm(self, wbName, filePath): #self.panelInput.get(),
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
        #     else:
        #         self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')
        # else:
        #     self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')

root = Tk()
root.title("PID data automation")
root.tk.call('wm', 'iconphoto', root._w, PhotoImage(file='Monocrystalline-solar-cell.png'))
app = Application(master=root)
app.mainloop()