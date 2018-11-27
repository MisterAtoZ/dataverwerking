import tkinter as tk
from tkinter import filedialog
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

        #set variables if config file excists
        try:
            with open('config.txt', 'r') as f:
                data = f.read()
                settings = data.split('|')
                print(str(settings))
                self.naamVar.set(settings[0])
                self.aantalVar.set(settings[1])
                self.F.set(settings[2])
                self.B.set(settings[3])
                self.R.set(settings[4])
                self.fileVar.set(settings[5])
        except IOError:
            pass

        #labels
        self.uurLabel = tk.Label(self, text='Hoeveel uren zijn de zonnepanelen gestressed?').grid(sticky='W',row=0)
        self.naamLabel = tk.Label(self, text='Welke zonnecellen zijn er getest geweest?').grid(sticky='W',row=1)
        self.aantalLabel = tk.Label(self, text='Hoeveel samples?').grid(sticky='W',row=1,column=2)
        self.kantLabel = tk.Label(self, text='Welke kanten?').grid(sticky='W', row=1,column=4)
        self.fileLabel = tk.Label(self, text='Kies waar de resultaten moeten in worden opgeslagen').grid(sticky='W',row=2)
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
        self.stopBtn = tk.Button(self, text='Stop', command=quit).grid(sticky='W',row=3, pady=4, padx=5)
        self.beginBtn = tk.Button(self, text='Begin', command=self.begin).grid(sticky='W',row=3, pady=4, padx=112)
        self.fileBtn = tk.Button(self, text='Pick file', command=self.pickFile).grid(sticky='W',row=3, pady=4, padx=50)

    def pickFile(self):
        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if self.filename != None:
            self.fileInput.insert(0,self.filename)

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

        if self.uurInput.get() =="":
            self.errorLabel.config(text='Error : Vul een uur in')
        elif self.naamInput.get() == "":
            self.errorLabel.config(text='Error : Vul een naam in')
        elif self.aantalInput.get() == "":
            self.errorLabel.config(text='Error : Vul een aantal samples in')
        elif self.F.get() == 0 & self.B.get() == 0 & self.R.get() == 0 :
            self.errorLabel.config(text='Error : Kies ten minste 1 kant')
        elif self.uurInput.get().isdigit():
            print('uur ' + str(self.uurInput.get()))
            if Main.Main.begin(self, self.naamInput.get(), self.uurInput.get(), self.aantalInput.get(), kanten, wbName, pad):
                self.errorLabel.config(text='Het bestand is opgeslagen')
                with open('config.txt', 'w') as f:
                    # naam
                    f.write(self.naamInput.get())
                    f.write('|')
                    # aantal
                    f.write(str(self.aantalInput.get()))
                    f.write('|')
                    # kanten
                    f.write(str(self.F.get()))
                    f.write('|')
                    f.write(str(self.B.get()))
                    f.write('|')
                    f.write(str(self.R.get()))
                    f.write('|')
                    # filedir
                    f.write(self.fileInput.get())
            else:
                self.errorLabel.config(text='Er is iets fout gelopen')
        else:
            self.errorLabel.config(text='Error : Vul een getal bij uur in')



root = tk.Tk()
app = Application(master=root)
app.mainloop()

#ingevuld formulier kunnen opslaan