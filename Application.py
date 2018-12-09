import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import Main

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        #variables
        self.uurVar = tk.IntVar()
        self.naamVar = tk.StringVar()
        self.aantalVar = tk.IntVar()
        self.F = tk.IntVar()
        self.B = tk.IntVar()
        self.R = tk.IntVar()
        self.fileVar = tk.StringVar()
        self.comboVar = tk.StringVar()
        self.comboList = []
        self.config_content = []
        self.samples = []

        #set variables if config file excists
        try:
            self.read_file()
        except IOError:
            self.comboList.append('')
            pass

        #labels
        self.uurLabel = tk.Label(self, text='Maximum hour of stressing?').grid(sticky='W',row=0)
        self.naamLabel = tk.Label(self, text='PV-cells names').grid(sticky='W',row=1)
        self.aantalLabel = tk.Label(self, text='Number of samples').grid(sticky='W',row=1,column=2)
        self.kantLabel = tk.Label(self, text='Sides').grid(sticky='W', row=1,column=4)
        self.fileLabel = tk.Label(self, text='Base file').grid(sticky='W',row=2)
        self.errorLabel = tk.Label(self, text='', fg='red')
        self.errorLabel.grid(sticky='W', row=4)

        #entries
        self.uurInput = tk.Entry(self, textvariable=self.uurVar)
        self.uurInput.grid(sticky='W',row=0, column=1)
        self.naamInput = tk.Entry(self, textvariable=self.naamVar)
        self.naamInput.grid(sticky='W',row=1, column=1)
        self.aantalInput = tk.Entry(self, textvariable=self.aantalVar)
        self.aantalInput.grid(sticky='W', row=1, column=3)
        self.fileInput = tk.Entry(self,textvariable=self.fileVar)
        self.fileInput.grid(sticky='WE',row=2,column=1,columnspan=7)

        #checkboxes
        self.kantF = tk.Checkbutton(self, text="F", variable=self.F, onvalue=1, offvalue=0).grid(row=1,column=5,sticky='W')
        self.kantB = tk.Checkbutton(self, text="B", variable=self.B, onvalue=1, offvalue=0).grid(row=1,column=6,sticky='W')
        self.kantR = tk.Checkbutton(self, text="R", variable=self.R, onvalue=1, offvalue=0).grid(row=1,column=7,sticky='W')

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
        i = int(self.comboVar.get().split(' - ')[0])
        with open('config.txt','r') as file:
            filedata = file.read()
        filedata= filedata.replace(self.config_content[i] + '<<>>', '')
        with open('config.txt', 'w') as file:
            file.write(filedata)

    def pickFile(self):
        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.filename != None:
            self.fileVar.set(self.filename)

    def select(self, event):
        print(self.comboVar.get())
        i = int(self.comboVar.get().split(' - ')[0])
        self.set_vars(i)

    def read_file(self):
        with open('config.txt', 'r') as f:
            data = f.read()
            if data != '':
                self.samples = data.split('<<>>')
                for i in range(0, len(self.samples) - 1, 1):
                    self.config_content.append(self.samples[i])
                    self.set_vars(i)
                    self.comboVar.set(str(i) + ' - ' + self.naamVar.get() + '_' + str(self.aantalVar.get()))
                    self.comboList.append(self.comboVar.get())
            else:
                self.comboList.append('')

    def set_vars(self, i):
        settings = self.samples[i].split('|')
        print(str(settings))
        self.naamVar.set(settings[0])
        self.aantalVar.set(settings[1])
        self.F.set(settings[2])
        self.B.set(settings[3])
        self.R.set(settings[4])
        self.fileVar.set(settings[5])

    def begin(self):
        #kanten
        kanten = []
        if self.F.get() != 0:
            kanten.append('F')
        if self.B.get() != 0:
            kanten.append('B')
        if self.R.get() != 0:
            kanten.append('R')
        print('kanten : ' + str(kanten))

        #wbName
        pad=''
        filenameSplit = self.fileInput.get().split('/')
        for i in range(0,len(filenameSplit)-1,1):
            pad = pad + str(filenameSplit[i]) + '/'
        wbName = filenameSplit[len(filenameSplit)-1]

        if self.uurInput.get() == "":
            self.errorLabel.config(text='Error : Missing hour')
        elif self.naamInput.get() == "":
            self.errorLabel.config(text='Error : Missing name')
        elif self.aantalInput.get() == "":
            self.errorLabel.config(text='Error : Missing number of samples')
        elif self.F.get() == 0 and self.B.get() == 0 and self.R.get() == 0 :
            self.errorLabel.config(text='Error : Missing sides')
        elif self.uurInput.get().isdigit():
            print('uur ' + str(self.uurInput.get()))
            if Main.Main.begin(self, self.naamInput.get(), self.uurInput.get(), self.aantalInput.get(), kanten, wbName, pad):
                self.errorLabel.config(text='File saved')
                config_text = self.naamInput.get() + '|' + str(self.aantalInput.get()) + '|' + str(self.F.get()) + '|' +\
                              str(self.B.get()) + '|' + str(self.R.get()) + '|' + self.fileInput.get()

                in_file = False
                for i in range(0,len(self.config_content),1):
                    if config_text == self.config_content[i]:
                        in_file = True

                if in_file == False:
                    with open('config.txt','a+') as f:
                        f.write(config_text)
                        f.write('<<>>')
            else:
                self.errorLabel.config(text='Something went wrong')
        else:
            self.errorLabel.config(text='Error : Hour must be a number')

root = tk.Tk()
app = Application(master=root)
app.mainloop()