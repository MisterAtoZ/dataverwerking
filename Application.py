import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import csv
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
        #variables
        self.hourVar = tk.IntVar()
        self.nameVar = tk.StringVar()
        self.amountVar = tk.IntVar()
        self.F = tk.IntVar()
        self.B = tk.IntVar()
        self.R = tk.IntVar()
        self.fileVar = tk.StringVar()
        self.comboVar = tk.StringVar()
        self.comboList = []
        self.configContent = []
        self.samples = []

        #set variables if config file excists
        try:
            self.readFile()
        except IOError:
            self.comboList.append('')
            pass

        #labels
        self.hourLabel = tk.Label(self, text='Maximum  duration of stressing? [h]').grid(sticky='W', row=0)
        self.nameLabel = tk.Label(self, text='PV-cells names').grid(sticky='W', row=1)
        self.amountLabel = tk.Label(self, text='Number of samples').grid(sticky='W', row=1, column=2)
        self.sideLabel = tk.Label(self, text='Sides').grid(sticky='W', row=1, column=4)
        self.fileLabel = tk.Label(self, text='Base file').grid(sticky='W',row=2)
        self.errorLabel = tk.Label(self, text='', fg='red')
        self.errorLabel.grid(sticky='W', row=4)

        #entries
        self.hourInput = tk.Entry(self, textvariable=self.hourVar)
        self.hourInput.grid(sticky='W', row=0, column=1)
        self.nameInput = tk.Entry(self, textvariable=self.nameVar)
        self.nameInput.grid(sticky='W', row=1, column=1)
        self.amountInput = tk.Entry(self, textvariable=self.amountVar)
        self.amountInput.grid(sticky='W', row=1, column=3)
        self.fileInput = tk.Entry(self,textvariable=self.fileVar)
        self.fileInput.grid(sticky='WE',row=2,column=1,columnspan=7)

        #checkboxes
        self.sideF = tk.Checkbutton(self, text="F", variable=self.F, onvalue=1, offvalue=0).grid(row=1, column=5, sticky='W')
        self.sideB = tk.Checkbutton(self, text="B", variable=self.B, onvalue=1, offvalue=0).grid(row=1, column=6, sticky='W')
        self.sideR = tk.Checkbutton(self, text="R", variable=self.R, onvalue=1, offvalue=0).grid(row=1, column=7, sticky='W')

        #buttons
        #self.stopBtn = tk.Button(self, text='Stop', command=quit).grid(sticky='W',row=3, pady=4, padx=5)
        self.beginBtn = tk.Button(self, text='Begin', command=self.begin).grid(sticky='W',row=3,column=2, pady=4, padx=20)
        self.fileBtn = tk.Button(self, text='Pick file', command=self.pickFile).grid(sticky='W',row=3, column=1, pady=4)
        self.rmBtn = tk.Button(self, text='Remove config', command=self.remove).grid(sticky='W',row=3, column=1,pady=4, padx=(70,0))

        #combobox
        self.combo = ttk.Combobox(self,textvariable=self.comboVar,values=self.comboList)
        self.combo.grid(row=0,column=3)
        self.combo.current(len(self.comboList)-1)
        self.combo.bind('<<ComboboxSelected>>', self.select)  # binding of user selection with a custom callback


    def remove(self):
        """
        removes a configuration form the config.txt file
        is called by rmBtn
        """
        i = int(self.comboVar.get().split(' - ')[0])
        with open('config.txt','r') as file:
            filedata = file.read()
        filedata= filedata.replace(self.configContent[i] + '<<>>', '')
        with open('config.txt', 'w') as file:
            file.write(filedata)


    def pickFile(self):
        """
        select a .xlsx file via the filedialog
        is called by fileBtn
        """
        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.filename != None:
            self.fileVar.set(self.filename)


    def select(self, event):
        """
        select a configuration with the comboBox and set the configuration variables in the gui
        """
        print(self.comboVar.get())
        i = int(self.comboVar.get().split(' - ')[0])
        self.setVars(i)


    def readFile(self):
        """
        read the config.txt file and put the configurations in the comboBox and in the gui
        """
        # with open ('config.csv', 'r') as f:
        #     data = f.read()
        #     reader = csv.reader(f)
        #     print(reader)
        #     print('data:' + data + '-----')
        #     if data != '':
        #         print('in if')
        #         lineCount=0
        #         for row in reader:
        #             print('row:' + row)
        #             if lineCount == 0:
        #                 print('0:' + row)
        #                 lineCount+=1
        #             else:
        #                 print(row + ';')
        #                 if row != '':
        #
        #                     #self.samples.append(row)
        #                     self.configContent.append(row)
        #                     self.setVars(lineCount-1)
        #                     self.comboVar.set(str(lineCount-1)+' - '+self.nameVar.get()+'_'+str(self.amountVar.get()))
        #                     self.comboList.append(self.comboVar.get())
        #                     lineCount+=1
        #     else:
        #         self.comboList.append('')
        with open('config.txt', 'r') as f:
            data = f.read()
            if data != '':
                self.samples = data.split('<<>>')
                for i in range(0, len(self.samples) - 1, 1):
                    self.configContent.append(self.samples[i])
                    self.setVars(i)
                    self.comboVar.set(str(i) + ' - ' + self.nameVar.get() + '_' + str(self.amountVar.get()))
                    self.comboList.append(self.comboVar.get())
            else:
                self.comboList.append('')


    def setVars(self, i):
        """
        set the variables from the configuration from the comboBox in the gui
        :param i: the number of the configuration
        """
        settings = self.samples[i].split('|') #configContent[i].split(',')
        print(str(settings))
        self.nameVar.set(settings[0])
        self.amountVar.set(settings[1])
        self.F.set(settings[2])
        self.B.set(settings[3])
        self.R.set(settings[4])
        self.fileVar.set(settings[5])


    def begin(self):
        """
        check if all variables are set correctly and starts the begin function of the Main class
        is called by beginBtn
        """
        #sides
        sides = []
        if self.F.get() != 0:
            sides.append('F')
        if self.B.get() != 0:
            sides.append('B')
        if self.R.get() != 0:
            sides.append('R')
        print('sides : ' + str(sides))

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
            self.errorLabel.config(text='Error : Missing hour')
        elif self.nameInput.get() == "":
            self.errorLabel.config(text='Error : Missing name')
        elif self.amountInput.get() == "":
            self.errorLabel.config(text='Error : Missing number of samples')
        elif self.hourInput.get().isdigit():
            print('uur ' + str(self.hourInput.get()))
            # if begin function is done : show 'file saved' and save the configuration if not already in config.txt
            if Main.Main.begin(self, self.nameInput.get(), self.hourInput.get(), self.amountInput.get(), sides, wbName, filePath):
                self.errorLabel.config(text='File saved')
                configText = self.nameInput.get() + '|' + str(self.amountInput.get()) + '|' + str(self.F.get()) + '|' + \
                              str(self.B.get()) + '|' + str(self.R.get()) + '|' + self.fileInput.get()
                inFile = False
                for i in range(0, len(self.configContent), 1):
                    if configText == self.configContent[i]:
                        inFile = True

                if inFile == False:
                    # with open('config.csv', mode='a+') as config:
                    #     fileWriter = csv.writer(config, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    #     fileWriter.writerow([self.nameInput.get(), str(self.amountInput.get()), str(self.F.get()), \
                    #                          str(self.B.get()), str(self.R.get()), self.fileInput.get()])
                    with open('config.txt','a+') as f:
                        f.write(configText)
                        f.write('<<>>')
            else:
                self.errorLabel.config(text='Something went wrong')
        else:
            self.errorLabel.config(text='Error : Hour must be a number')

root = tk.Tk()
app = Application(master=root)
app.mainloop()