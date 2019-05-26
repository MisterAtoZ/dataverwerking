import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
from PIL import ImageTk, Image
import glob
import openpyxl
import os
from os import listdir
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
#matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib.figure import Figure
import Data
import AlgemeneInfo
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
        gui_style = ttk.Style()
        #gui_style.configure('My.TButton', background='#2E9CCA')
        gui_style.map("My.TButton",
    #foreground=[('pressed', 'red'), ('active', bgColor)],
    background=[('pressed', '!disabled', bgColor), ('active', bgColor), ('!active', bgColor)]
    )
        gui_style.configure('My.TFrame', background=bgColor)
        gui_style.configure('My.TNotebook', background=fgColor)
        gui_style.configure('My.TLabel', background=bgColor, foreground=fgColor, font=('verdana', 10, ''))
        gui_style.configure('My.TCheckbutton', background=bgColor, foreground=fgColor, font=('verdana', 10, ''))
        gui_style.configure('My.TLabelframe.Label', background=bgColor, foreground=fgColor)
        gui_style.configure('My.TRadiobutton', background=bgColor, foreground=fgColor, font=('verdana', 10, ''))
        gui_style.configure('My.TSeparator', background='#000000')
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

        self.tabBifi = ttk.Frame(tabControl, style='My.TFrame')
        tabControl.add(self.tabBifi, text='Processing')

        # tabPsc = ttk.Frame(tabControl, style='My.TFrame')
        # tabControl.add(tabPsc, text='Thin film')

        self.tabInfo = ttk.Frame(tabControl, style='My.TFrame')
        tabControl.add(self.tabInfo, text='Info')

        #variables
        self.f = Figure(figsize=(4, 4), dpi=100)
        self.ax = self.f.add_subplot(111)
        self.fb = Figure(figsize=(4, 4), dpi=100)
        self.bx = self.fb.add_subplot(111)
        self.filePath = ''
        self.samples = []
        self.wbName = ''
        self.hours = []
        self.hourRb = tk.IntVar()
        self.machineRb = tk.IntVar()
        # self.sampleRb = []
        self.sampleRb = tk.IntVar()
        self.hourVar = tk.IntVar()
        self.hourVar2 = tk.IntVar()
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

        self.fileSm = tk.StringVar()
        self.comboSm = tk.StringVar()
        self.comboListSm = [''] # nr + configContent
        self.configSm = []      # filepath

        #set variables if config file excists
        # try:
        #     self.readFile()
        # except IOError:
        #     self.comboList.append('') # voor Psc en Sm ook -----------------------------
        #     pass

        #Processing
            # Machine choise
        ttk.Separator(self.tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=0, columnspan=6, sticky='WE', padx=5)
        self.machineLabel = ttk.Label(self.tabBifi, text='Select machine', style='My.TLabel').grid(sticky='W', row=0, column=0, columnspan=2, padx=10)
        self.RbLoana = ttk.Radiobutton(self.tabBifi, text='LOANA', variable=self.machineRb, value=1, style='My.TRadiobutton', command = lambda:self.selectMachine())\
            .grid(row=1,sticky='W', padx=5, pady=(0,5))
        self.RbThin = ttk.Radiobutton(self.tabBifi, text='Thin film', variable=self.machineRb, value=2, style='My.TRadiobutton', command=lambda: self.selectMachine())\
            .grid(row=1, column=1,sticky='W', pady=(0,5))
        self.machineRb.set(1)
        self.RbSwitch = ttk.Radiobutton(self.tabBifi, text='Switch matrix', variable=self.machineRb, value=3,
                                        style='My.TRadiobutton', command=lambda: self.selectMachine()) \
            .grid(row=1, column=2, sticky='W', pady=(0,5))
        self.machineRb.set(1)
        ttk.Separator(self.tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=2, columnspan=6, sticky='WE', padx=5)
            # Hour choise
        self.hourLabel = ttk.Label(self.tabBifi, text='Stressing duration?', style='My.TLabel').grid(sticky='W', row=2, column=0, columnspan=3, padx=10)
        self.hourRbAll = ttk.Radiobutton(self.tabBifi, text='All', variable=self.hourRb, value=1,style='My.TRadiobutton').grid(row=3, sticky='W', padx=5)
        self.hourRbEntry = ttk.Radiobutton(self.tabBifi, text='Between', variable=self.hourRb, value=2,style='My.TRadiobutton').grid(row=3, column=1,sticky='W')
        self.hourRb.set(1)
        self.hourInput = ttk.Entry(self.tabBifi, textvariable=self.hourVar)
        self.hourInput.grid(sticky='WE', row=3, column=2, pady=5,padx=(5,5))
        self.hourLabel = ttk.Label(self.tabBifi, text='and', style='My.TLabel').grid(sticky='W', row=3,column=3,padx=5)
        self.hourInput2 = ttk.Entry(self.tabBifi, textvariable=self.hourVar2)
        self.hourInput2.grid(sticky='WE', row=3, column=4, pady=5, padx=(5, 5))
        ttk.Separator(self.tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=4, columnspan=6, sticky='WE', padx=5)
            # Excel file
        self.excelLabel = ttk.Label(self.tabBifi, text='Select Excel file', style='My.TLabel').grid(sticky='W',row=4,column=0,columnspan=2, padx=10)
        #self.configLabel = ttk.Label(self.tabBifi, text='Select previous configuration', style='My.TLabel').grid(sticky='W',row=1,column=3)
        # self.combo = ttk.Combobox(self.tabBifi, textvariable=self.comboVar, values=self.comboList)
        # self.combo.grid(row=5, column=0, columnspan=3, sticky='WE', padx=5)
        # self.combo.current(len(self.comboList) - 1)
        # self.combo.bind('<<ComboboxSelected>>', self.select)
        # self.rmBtn = ttk.Button(self.tabBifi, text='Remove config', style='My.TButton')
        # self.rmBtn.bind('<ButtonRelease-1>', self.remove)
        # self.rmBtn.grid(sticky='WE', row=5, column=4, padx=5)
        self.fileInput = ttk.Entry(self.tabBifi, textvariable=self.fileVar)
        self.fileInput.grid(row=6, column=1, columnspan=5, sticky='WE', padx=5, pady=5)
        self.fileBtn = ttk.Button(self.tabBifi, text='Select file', style='My.TButton')
        self.fileBtn.bind('<ButtonRelease-1>', self.pickFile)
        self.fileBtn.grid(sticky='WE', row=6, column=0, padx=5, pady=5)
        # self.fileBtn.config(width=9)
            # Data choise
        ttk.Separator(self.tabBifi, orient=HORIZONTAL, style='My.TSeparator').grid(row=7, columnspan=6, sticky='WE', padx=5)
        self.dataLabel = ttk.Label(self.tabBifi, text='Select data types', style='My.TLabel').grid(sticky='W', row=7, column=0,columnspan=2, padx=10)
        self.iv = ttk.Checkbutton(self.tabBifi, text="IV", variable=self.ivCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.iv.grid(row=8, column=0, sticky='W', padx=5, pady=5)
        self.eqe = ttk.Checkbutton(self.tabBifi, text="EQE", variable=self.eqeCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.eqe.grid(row=8, column=1, sticky='W', pady=5)
        self.photo = ttk.Checkbutton(self.tabBifi, text="Photo", variable=self.photoCb, onvalue=1, offvalue=0,style='My.TCheckbutton')
        self.photo.grid(row=8, column=2, sticky='W', pady=5)
            # Start
        self.beginBtn = ttk.Button(self.tabBifi, text='Begin', command=self.begin, style='My.TButton')
        self.beginBtn.grid(sticky='W',row=9,column=0,padx=5,pady=5)
        # self.beginBtn.config(width=9)
        self.errorLabel = ttk.Label(self.tabBifi, text='', background=bgColor, foreground=red, font=('verdana', 10, ''))
        self.errorLabel.grid(sticky='W', row=9, column=1, columnspan=3)
        img = Image.open("UHasselt.png")
        basewidth = 120 #70 voor zelfde breedte
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
        img = img.resize((basewidth, hsize), Image.ANTIALIAS)
        img = ImageTk.PhotoImage(img)
        label1 = Label(self.tabBifi, image=img, background='white')
        label1.image = img
        label1.grid(row=0, column=5, columnspan=2, padx=5, pady=(5,5), sticky='E')#, columnspan=2) #pady=(0,5) voor breedte

        #configure column 2 to stretch with the window
        self.tabBifi.grid_columnconfigure(5, weight=1)

        #Info
        self.cbFrame = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.cbFrame.grid(row=0, columnspan=5, sticky='W', padx=5)
        self.btn = ttk.Button(self.tabInfo, text='Begin', command=self.graph, style='My.TButton')
        self.btn.grid(sticky='W', row=1, column=0, padx=5, pady=5)
        self.gFrame = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.gFrame.grid(row=2, rowspan=2, sticky='W', padx=5)
        self.gFrame2 = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.gFrame2.grid(row=2, rowspan=2, column=1, sticky='W', padx=5)
        self.tFrame = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.tFrame.grid(row=4, column=0, sticky='W', padx=5)
        self.pFrame = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.pFrame.grid(row=2, column=2, sticky='W', padx=5)

        self.canvas = FigureCanvasTkAgg(self.f, self.gFrame)
        self.canvas.show()
        self.canvas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2TkAgg(self.canvas, self.gFrame)
        self.toolbar.update()
        self.canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.canvas2 = FigureCanvasTkAgg(self.fb, self.gFrame2)
        self.canvas2.show()
        self.canvas2.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2TkAgg(self.canvas2, self.gFrame2)
        self.toolbar.update()
        self.canvas2._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # configure column 2 to stretch with the window
        # self.tabInfo.grid_columnconfigure(0, weight=1)
        # self.tabInfo.grid_rowconfigure(3, weight=1)

        # myscrollbar = Scrollbar(self.tabInfo, orient="vertical")
        # myscrollbar.grid(row=1, column=3, sticky='E', rowspan=100)

    def graph(self):
        self.ax.clear()
        self.bx.clear()
        iv = []
        print('----')
        print(self.sampleRb.get())
        # for i in range(0, len(self.sampleRb)):
        #     print(self.samples[i])
        #     print(self.sampleRb[i].get())
        #     if self.sampleRb[i].get() == 1:
        i = self.sampleRb.get()
        self.f = plt.figure(1)
        for h in self.hours:
            print(h)
            if self.machineRb.get() == 1:
                if os.path.exists(self.filePath + '/' + str(h) + '/' + self.samples[i] + '/IV/' + self.samples[i] + '.lgt'):
                    iv = Data.Data.getDataList(self.filePath + '/' + str(h) + '/' + self.samples[i] + '/IV/' + self.samples[i] + '.lgt')
                    div = Data.Data.getDataList(self.filePath + '/' + str(h) + '/' + self.samples[i] + '/IV/' + self.samples[i] + '.drk')
                    self.ax.plot(iv[1], iv[0], label=h + 'h')
                    self.bx.plot(div[1], div[0], label=h + 'h')
            else:
                if os.path.exists(self.filePath + '/' + self.samples[i] + '/' + h + '.csv'):
                    iv = Data.Data.getDataListSm(self.filePath + '/' + self.samples[i] + '/' + h + '.csv')
                    self.bx.plot(iv[1], iv[0], label=h + 'h')
        self.ax.set_xlabel('Voltage [V]')
        self.ax.set_ylabel('Current [A]')
        self.ax.set_title('Light IV')
        self.ax.legend()
        self.bx.set_xlabel('Voltage [V]')
        self.bx.set_ylabel('Current [A]')
        self.bx.set_title('Dark IV')
        self.bx.legend()
        self.canvas.draw()
        self.canvas2.draw()

        self.elImage()
        self.table()

    def table(self):
        print('table ---')
        for widget in self.tFrame.winfo_children():
            print(widget)
            widget.destroy()
        print(self.hours[-1])
        print(self.filePath + self.hours[-1] + '-' + self.wbName)
        dirname = self.filePath
        filespec = self.hours[-1] + '-' + '*' + self.wbName
        print(glob.glob(os.path.join(dirname,filespec)))
        file = glob.glob(os.path.join(dirname,filespec))[0]
        if os.path.exists(file):
            ttk.Separator(self.tFrame, orient=HORIZONTAL, style='My.TSeparator').grid(row=1, columnspan=9, sticky='WE', padx=5)

            columnH = 0
            columnP = 2
            columnI = 4
            columnV = 6
            columnF = 8

            self.titleH = ttk.Label(self.tFrame, text='Time [h]', style='My.TLabel').grid(sticky='NW', row=0, column=columnH, padx=5)
            ttk.Separator(self.tFrame, orient=VERTICAL, style='My.TSeparator').grid(row=0, column=1, rowspan=len(self.hours)+2, sticky='NS')
            self.titleP = ttk.Label(self.tFrame, text='%PID', style='My.TLabel').grid(sticky='NW', row=0, column=columnP, padx=5)
            ttk.Separator(self.tFrame, orient=VERTICAL, style='My.TSeparator').grid(row=0, column=3,rowspan=len(self.hours) + 2, sticky='NS')
            self.titleI = ttk.Label(self.tFrame, text='Isc [mA]', style='My.TLabel').grid(sticky='NW', row=0, column=columnI,padx=5)
            ttk.Separator(self.tFrame, orient=VERTICAL, style='My.TSeparator').grid(row=0, column=5,rowspan=len(self.hours) + 2, sticky='NS')
            self.titleV = ttk.Label(self.tFrame, text='Voc [mV]', style='My.TLabel').grid(sticky='NW', row=0, column=columnV,padx=5)
            ttk.Separator(self.tFrame, orient=VERTICAL, style='My.TSeparator').grid(row=0, column=7,rowspan=len(self.hours) + 2,sticky='NS')
            self.titleF = ttk.Label(self.tFrame, text='FF [%]', style='My.TLabel').grid(sticky='NW', row=0, column=columnF,padx=5)

            for i in range(0, len(self.hours)):
                # get %PID from Excel
                wb = openpyxl.load_workbook(file, data_only=True)
                j = self.sampleRb.get()
                activeSheet = wb[self.samples[j]]
                if activeSheet.cell(row=i+2, column=17).value is not None:
                    labelP = round(activeSheet.cell(row=i+2, column=17).value,2)
                else:
                    labelP = '' # label over label -> update
                if activeSheet.cell(row=i + 2, column=5).value is not None:
                    labelI = activeSheet.cell(row=i + 2, column=5).value
                else:
                    labelI = ''
                if activeSheet.cell(row=i + 2, column=7).value is not None:
                    labelV = activeSheet.cell(row=i + 2, column=7).value
                else:
                    labelV = ''
                if activeSheet.cell(row=i + 2, column=8).value is not None:
                    labelF = activeSheet.cell(row=i + 2, column=8).value
                else:
                    labelF = ''

                print(labelP)
                ttk.Label(self.tFrame, text=self.hours[i], style='My.TLabel').grid(row=i + 2, column=columnH, padx=5, sticky='WE')
                ttk.Label(self.tFrame, text=labelP, style='My.TLabel').grid(row=i+2, column=columnP, padx=5, sticky='WE')
                ttk.Label(self.tFrame, text=labelI, style='My.TLabel').grid(row=i + 2, column=columnI, padx=5, sticky='WE')
                ttk.Label(self.tFrame, text=labelV, style='My.TLabel').grid(row=i + 2, column=columnV, padx=5, sticky='WE')
                ttk.Label(self.tFrame, text=labelF, style='My.TLabel').grid(row=i + 2, column=columnF, padx=5, sticky='WE')


    def elImage(self):
        print(self.hours[-1])
        for widget in self.tFrame.winfo_children():
            print(widget)
            widget.destroy()
        name = self.wbName.split('.')[0] + '.jpg'
        file = self.filePath + self.hours[-1] + '-' + name
        print(file)
        if os.path.exists(file):
            img = Image.open(file)
            hsize = 400
            hpercent = (hsize/float(img.size[1]))
            basewidth = int((float(img.size[0]) * float(hpercent))) #1120
            if basewidth > 1000:
                basewidth = 1000  # 70 voor zelfde breedte
                wpercent = (basewidth / float(img.size[0]))
                hsize = int((float(img.size[1]) * float(wpercent)))
            img = img.resize((basewidth, hsize), Image.ANTIALIAS)
            img = ImageTk.PhotoImage(img)
            label1 = Label(self.pFrame, image=img, background='white')
            label1.image = img
            label1.grid(row=0, column=0, padx=5, pady=(5, 5),
                        sticky='W')  # , columnspan=2) #pady=(0,5) voor breedte


    def selectMachine(self):
        if self.machineRb.get() == 1:
            self.iv.config(state=NORMAL)
            self.eqe.config(state=NORMAL)
            self.photo.config(state=NORMAL)
        else:
            self.iv.config(state=DISABLED)
            self.eqe.config(state=DISABLED)
            self.photo.config(state=DISABLED)

    # def remove(self, event):
    #     """
    #     removes a configuration form the config.txt file
    #     is called by rmBtn
    #     """
    #     caller = event.widget
    #     frame = self.getFrame(caller)
    #
    #     if frame == 'Bifi':
    #         i = int(self.comboVar.get().split(' - ')[0]) - 1
    #         #filepath = self.comboList[i].split('|')[1]
    #         filepath = self.configContent[i].split('-')[1]
    #         print(filepath)
    #         with open('config.txt','r') as file:
    #             filedata = file.read()
    #         filedata= filedata.replace('Bifi|' + filepath + '\n', '')
    #         with open('config.txt', 'w') as file:
    #             file.write(filedata)
    #     # if frame == 'Psc':
    #     #     i = int(self.comboPsc.get().split(' - ')[0]) - 1
    #     #     with open('config.txt', 'r') as file:
    #     #         filedata = file.read()
    #     #     filedata = filedata.replace('Psc|' + self.configPsc[i] + '\n', '')
    #     #     with open('config.txt', 'w') as file:
    #     #         file.write(filedata)

    def pickFile(self, event):
        """
        select a .xlsx file via the filedialog
        is called by fileBtn
        """
        caller = event.widget
        print(caller)
        self.cbFrame.destroy()
        self.cbFrame = None
        self.cbFrame = ttk.Frame(self.tabInfo, style='My.TFrame')
        self.cbFrame.grid(row=0, columnspan=5, sticky='W', padx=5)
        self.filePath = ''
        self.samples = []
        self.wbName = ''
        self.hours = []
        self.sampleRb = tk.IntVar()

        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.filename != None:
                self.fileVar.set(self.filename)
                self.wbName = self.filename.split('/')[-1]
                print(self.wbName)
                self.filePath = self.filename.replace(self.wbName, '')
                # pathsplit = self.filename.split('/')[:-1]
                # self.filePath = ''
                # for i in pathsplit:
                #     self.filePath = self.filePath + i + '/'
                print(self.filePath)
                for f in listdir(self.filePath):
                    if self.machineRb.get() == 1:
                        if os.path.isdir(self.filePath + f):
                            for g in listdir(self.filePath + f):
                                if os.path.isdir(self.filePath + f + '/' + g):
                                    self.samples.append(g)
                                    rb = tk.IntVar()
                                    # self.sampleRb.append(rb)
                        break
                    elif self.machineRb.get() == 2:
                        print('thin film')
                    else:
                        if os.path.isdir(self.filePath + f):
                            self.samples.append(f)
                            rb = tk.IntVar()
                            # self.sampleRb.append(rb)

                print(self.samples)
                if self.samples != []:
                    if self.machineRb.get() == 1:
                        # wb = openpyxl.load_workbook(self.filename, data_only=True)
                        # self.hours = AlgemeneInfo.AlgemeneInfo.calculateHours(self, wb)
                        # for i in range(0, len(self.hours), 1):
                        #     self.hours[i] = int(round(self.hours[i], 0))
                        subfolders = [f.name for f in os.scandir(self.filePath) if f.is_dir()]
                        subfolders.sort(key=float)
                        self.hours = subfolders
                    elif self.machineRb.get() == 3:
                        print(self.samples[0])
                        modulePath = self.filePath + self.samples[0]
                        files = [f.name for f in os.scandir(modulePath)]
                        for f in files:
                            if f.endswith('.csv'):
                                self.hours.append(f[:-4])

                    print('add checkboxes')

                    for i in range(0,len(self.samples)):
                        rb = ttk.Radiobutton(self.cbFrame, text=self.samples[i], variable=self.sampleRb, value=i,
                                        style='My.TRadiobutton')
                        # cb = ttk.Checkbutton(self.cbFrame, text=self.samples[i], variable=self.sampleRb[i], style='My.TCheckbutton', )
                        rb.pack(side='left', fill=None, expand=False, padx=(0,5))

            # if frame == 'Psc':
            #     self.filePsc.set(self.filename)
            # if frame == 'Sm':
            #     self.fileSm.set(self.filename)


    # def select(self, event):
    #     """
    #     select a configuration with the comboBox and set the configuration variables in the gui
    #     """
    #     if self.machineRb.get() == 1:
    #         print(self.comboVar.get())
    #         nr = int(self.comboVar.get().split(' - ')[0])-1
    #         path = self.configContent[nr]
    #         self.setVars('Bifi', path)
    #     if self.machineRb.get() == 2:
    #         print(self.comboPsc.get())
    #         path = self.comboPsc.get().split(' - ')[1]
    #         self.setVars('Psc', path)
    #     if self.machineRb.get() == 3:
    #         print(self.comboSm.get())
    #         path = self.comboSm.get().split(' - ')[1]
    #         self.setVars('Sm', path)
    #
    # def readFile(self):
    #     """
    #     read the config.txt file and put the configurations in the comboBox and in the gui
    #     """
    #     with open('config.txt', 'r') as f:
    #         data = f.read()
    #         if data != '':
    #             self.sampels = data.split('\n')
    #             for i in range(0, len(self.sampels) - 1, 1):
    #                 frame = self.sampels[i].split('|')[0]
    #                 path = self.sampels[i].split('|')[1]
    #                 nr = len(self.configContent) + 1
    #                 self.configContent.append(str(nr) + '-' + path)
    #                 self.setVars(frame, path)
    #                 filenames = self.fileVar.get().split('/')
    #                 filename = filenames[len(filenames)-1]
    #                 self.comboVar.set(str(nr) + ' - ' + filename)
    #                 self.comboList.append(str(self.comboVar.get()))
    #                 # self.comboList.append(self.comboVar.get())
    #                 # if frame == 'Psc':
    #                 #     self.configPsc.append(path)
    #                 #     self.setVars(frame, path)
    #                 #     self.comboPsc.set(str(len(self.configPsc)) + ' - ' + self.filePsc.get())
    #                 #     self.comboListPsc.append(self.comboPsc.get())
    #         else:
    #             self.comboList.append('')
    #         print(self.comboList)
    #         print(self.comboVar.get())
    #
    #
    # def setVars(self, frame, path):
    #     """
    #     set the variables from the configuration from the comboBox in the gui
    #     :param i: the number of the configuration
    #     """
    #     self.fileVar.set(path)
    #     # if frame == 'Psc':
    #     #     self.filePsc.set(path)
    #     # if frame == 'Sm':
    #     #     self.filePsc.set(path)

    def begin(self):
        rb = self.machineRb.get()
        if rb == 1:
            self.beginBifi()
        elif rb == 2:
            self.beginPsc()
        elif rb == 3:
            self.beginSm()
        else:
            self.errorLabel.config(text='Error with machine selection')

    def beginBifi(self):
        """
        check if all variables are set correctly and starts the beginBifi function of the Main class
        is called by beginBtn
        """
        print('beginBifi')
        #wbName
        # self.filePath=''
        # filenameSplit = self.fileInput.get().split('/')
        # for i in range(0,len(filenameSplit)-1,1):
        #     self.filePath = self.filePath + str(filenameSplit[i]) + '/'
        # wbName = filenameSplit[-1]
        # check if all variables are set correctly
        # yes: begin
        # no: show error message
        print('---')
        print(self.hourInput.get())
        print(self.hourInput2.get())
        print(self.hourRb.get())
        if self.hourRb.get() == 2 and(self.hourInput.get() == '' or self.hourInput2.get() == ''):
            self.errorLabel.config(text='Error : No hour')
        elif self.hourRb.get() == 2 and (not self.hourInput.get().isdigit() or not self.hourInput2.get().isdigit()):
            self.errorLabel.config(text='Error : Hour must be a number')
        elif self.hourRb.get() == 2 and (int(self.hourInput.get()) >= int(self.hourInput2.get())):
            self.errorLabel.config(text='Error : Hour interval is not correct')
        else:
            print('uur ' + str(self.hourInput.get()))
            # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
            if(self.fileInput.get().endswith('.xlsx')):
                self.errorLabel.config(text='Running...')
                self.errorLabel.update()
                if self.hourRb.get() == 1:
                    hours = -1
                else:
                    hours = self.hourInput.get() #+ '-' + self.hourInput2.get()
                if Main.Main.beginBifi(self, hours, self.wbName, self.filePath, self.ivCb.get(), self.eqeCb.get(), self.photoCb.get()):
                    self.errorLabel.config(text='File saved')

                    # configText = 'Bifi|' + self.fileInput.get()
                    #
                    # # check if the file is already in config.txt and add it if not
                    # inFile = False
                    # for i in range(0, len(self.configContent), 1):
                    #     if configText == 'Bifi|' + self.configContent[i]:
                    #         inFile = True
                    #
                    # if inFile == False:
                    #     with open('config.txt','a+') as f:
                    #         f.write(configText)
                    #         f.write('\n')
                    #         self.configContent.append(filePath + wbName)
                else:
                    self.errorLabel.config(text='Error: Something went wrong')
            else:
                self.errorLabel.config(text='Error : Not a .xlsx file')

    def beginPsc(self):
        """
        check if all variables are set correctly and starts the beginPsc function of the Main class
        is called by beginBtnPsc
        """
        print('beginPsc')
        #wbName
        # self.filePath=''
        # filenameSplit = self.fileInputPsc.get().split('/')
        # for i in range(0,len(filenameSplit)-1,1):
        #     self.filePath = self.filePath + str(filenameSplit[i]) + '/'
        # wbName = filenameSplit[-1]
        # check if all variables are set correctly
        # yes: begin
        # no: show error message

        # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
        if(self.fileInput.get().endswith('.xlsx')):
            self.errorLabel.config(text='Running...')
            self.errorLabel.update()
            if Main.Main.beginPsc(self, self.wbName, self.filePath):
                self.errorLabel.config(text='File saved')

                # configText = 'Psc|' + self.fileInputPsc.get()
                #
                # # check if the file is already in config.txt and add it if not
                # inFile = False
                # for i in range(0, len(self.configPsc), 1):
                #     if configText == 'Psc|' + self.configPsc[i]:
                #         inFile = True
                #
                # if inFile == False:
                #     with open('config.txt','a+') as f:
                #         f.write(configText)
                #         f.write('\n')
                #         self.configPsc.append(self.filePath + self.wbName)
            else:
                self.errorLabel.config(text='Error: Something went wrong')
        else:
            self.errorLabel.config(text='Error : Not a .xlsx file')

    def beginSm(self):
        """
        check if all variables are set correctly and starts the beginSm function of the Main class
        is called by beginBtnSm
        """
        print('beginSw')
        #wbName
        # self.filePath=''
        # filenameSplit = self.fileInput.get().split('/')
        # for i in range(0,len(filenameSplit)-1,1):
        #     self.filePath = self.filePath + str(filenameSplit[i]) + '/'
        # wbName = filenameSplit[len(filenameSplit)-1]



        # check if all variables are set correctly
        # yes: begin
        # no: show error message

        # if self.panelInput.get() == "":
        #     self.errorLabelSm.config(text='Error : No panel')
        # elif self.panelInput.get().isdigit():
        #     if int(self.panelInput.get()) > 0 and int(self.panelInput.get()) <= 6:
        #         # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
        if(self.fileInput.get().endswith('.xlsx')):
            self.errorLabel.config(text='Running...')
            self.errorLabel.update()
            if Main.Main.beginSm(self, self.wbName, self.filePath): #self.panelInput.get(),
                self.errorLabel.config(text='File saved')

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
                self.errorLabel.config(text='Error: Something went wrong')
        else:
            self.errorLabel.config(text='Error : Not a .xlsx file')
        #     else:
        #         self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')
        # else:
        #     self.errorLabelSm.config(text='Error : Panel must be a number in the range 1 to 6')

root = Tk()
root.title("PID data automation")
root.tk.call('wm', 'iconphoto', root._w, PhotoImage(file='Monocrystalline-solar-cell.png'))
app = Application(master=root)
app.mainloop()