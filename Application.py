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
        self.uurLabel = tk.Label(self, text='Hoeveel uren zijn de zonnepanelen gestressed?').grid(sticky='W',row=0)
        self.naamLabel = tk.Label(self, text='Welke zonnecellen zijn er getest geweest?').grid(sticky='W',row=1)
        self.errorLabel = tk.Label(self, text='', fg='red')
        self.errorLabel.grid(sticky='W', row=4)

        uurVar = tk.IntVar()
        self.uurInput = tk.Entry(self, textvariable=uurVar)
        self.uurInput.grid(sticky='W',row=0, column=1)
        naamVar = tk.StringVar()
        self.naamInput = tk.Entry(self, textvariable=naamVar)
        self.naamInput.grid(sticky='W',row=1, column=1)

        self.stopBtn = tk.Button(self, text='Stop', command=quit).grid(sticky='W',row=3, pady=4, padx=5)
        self.beginBtn = tk.Button(self, text='Begin', command=self.begin).grid(sticky='W',row=3, pady=4, padx=112)
        self.fileBtn = tk.Button(self, text='Pick file', command=self.pickFile).grid(sticky='W',row=3, pady=4, padx=50)

    def pickFile(self):
        self.filename = tk.filedialog.askopenfilename(initialdir="/", title="Select file",filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))

    def begin(self):
        pad=''
        filenameSplit = self.filename.split('/')
        for i in range(0,len(filenameSplit)-1,1):
            pad = pad + str(filenameSplit[i]) + '/'
        wbName = filenameSplit[len(filenameSplit)-1]
        print(filenameSplit)
        print(pad)
        print(wbName)
        if self.uurInput.get() =="":
            self.errorLabel.config(text='Error : Vul een uur in')
        elif self.naamInput.get() == "":
            self.errorLabel.config(text='Error : Vul een naam in')
        elif self.uurInput.get().isdigit():
            if Main.Main.begin(self, self.naamInput.get(), self.uurInput.get(), wbName, pad):
                self.errorLabel.config(text='Het bestand is opgeslagen')
            else:
                self.errorLabel.config(text='Er is iets fout gelopen')
        else:
            self.errorLabel.config(text='Error : Vul een getal bij uur in')


root = tk.Tk()
app = Application(master=root)
app.mainloop()

#ingevuld formulier kunnen opslaan