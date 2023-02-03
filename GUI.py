from tkinter import *
from tkinter import messagebox
import tkinter.ttk as tk
from tkinter.filedialog import *
from tkinter.messagebox import *
from start import *


def help_window():
    showinfo(title="Hilfe", message="""Bei Problemen melden sie sich bei der folgende e-mail
e-mail: eren.aksu@nationalexpress.de""")


class GUI:

    def programmSchliessen(self):
        self.master.destroy()

    def __init__(self, master):
        self.datei = None
        self.speicherort = None

        self.master = master
        self.vs = StringVar()
        self.vs1 = StringVar()

        self.paddings = {'padx': 5, 'pady': 5}

        self.options = ['Betriebstage-Tool', 'PDF Cutter']
        self.master.title("National Express Technik-Tool")

        self.menubar = Menu(self.master)

        self.close_menu = Menu(self.menubar)
        self.close_menu.add_command()

        self.menubar = Menu(self.master)

        self.close_menu = Menu(self.menubar)
        self.close_menu.add_command(label="Programm schließen", command=self.programmSchliessen)

        self.help_menu = Menu(self.menubar)

        self.menubar.add_cascade(label="Programm", menu=self.close_menu)
        self.menubar.add_cascade(label="Hilfe", command=help_window)

        self.master.config(menu=self.menubar)

        self.master.geometry("500x500")
        self.master.minsize(500, 500)
        self.master.maxsize(500, 500)
        self.master.resizable(False, False)

        self.bt_dateiWaehlen = tk.Button(self.master, text="Datei Wählen", command=self.getPath)
        self.bt_speicheortWaehlen = tk.Button(self.master, text="Speicherort Wählen", command=self.getSavePath)
        self.bt_start = tk.Button(self.master, text="Programm Starten", command=self.start)

        self.bt_dateiWaehlen.place(anchor=CENTER, relx=.5, rely=.2)
        self.bt_speicheortWaehlen.place(anchor=CENTER, relx=.5, rely=.5)
        self.bt_start.place(anchor=CENTER, relx=.5, rely=.8)

        self.run = Start()

    def start(self):
        self.run.start(self.datei, self.speicherort)

    def getPath(self):
        path = askopenfilename(title="Wählen sie bitte aus",
                               filetypes=[("Excel-Datei", "*.xlsx"), ("Alle Dateien", "*.*")])

        self.datei = path

    def getSavePath(self):
        path = askdirectory()
        self.speicherort = path

        if path == "":
            messagebox.showinfo(title="Keine Datei Ausgewählt!", message="Bitte wählen sie eine Datei aus")
