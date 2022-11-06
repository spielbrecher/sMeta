import os
import tkinter.ttk
from tkinter import Tk, LEFT, RIGHT, BOTH, RAISED, messagebox, Label, Entry, StringVar
from tkinter.ttk import Frame, Button, Style, Combobox
from tkinter.filedialog import askopenfilename

from pandastable import Table, TableModel
import pandas as pd

from estimate_frame import EstimateFrame


class MainFrame(Tk):

    def __init__(self):
        super().__init__()

        self.filename = None  # filename of estimate raw version

        # Setting the frame
        self.frame = Frame(self, borderwidth=1)
        self.title("sMeta")
        self.center_window()
        self.frame.pack(fill=BOTH, expand=True)

        # Choose window style
        self.style = Style()
        themes = self.style.theme_names()  # all available themes
        self.style.theme_use(themes[1])

        # interface

        # table
        self.df = pd.DataFrame()
        self.table = Table(self.frame, dataframe=self.df, showtoolbar=False, showstatusbar=False)
        self.table.show()

        # buttons
        self.closeButton = Button(self, text="Закрыть", command=self.quit)
        self.closeButton.pack(side=RIGHT, padx=5, pady=5)
        self.loadButton = Button(self, text="Разобрать", command=self.parse_estimate)
        self.loadButton.pack(side=RIGHT)
        self.loadButton = Button(self, text="Загрузить", command=self.load_estimate)
        self.loadButton.pack(side=RIGHT)

    # center the main frame
    def center_window(self):
        w = 1200
        h = 800
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def load_estimate(self):
        Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        # show an "Open" dialog box and return the path to the selected file
        self.filename = askopenfilename(initialdir=os.getcwd(),
                                   title="Открыть смету",
                                   filetypes=(("Excel", "*.xlsx"), ("all files", "*.*")))
        # read for Pandas Table at first and work with it by other libraries
        try:
            self.df = pd.read_excel(self.filename)
            self.table.model.df = self.df
            self.table.redraw()

        except:
            tkinter.messagebox.showerror("Ошибка", "w")

    def parse_estimate(self):
        estimate = EstimateFrame(self.df)

