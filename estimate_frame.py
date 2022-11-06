import os
import tkinter.ttk
from tkinter import Tk, LEFT, RIGHT, BOTH, RAISED, messagebox, Label, Entry, StringVar
from tkinter.ttk import Frame, Button, Style, Combobox
import tkinter.filedialog as dialog

from pandastable import Table, TableModel
import pandas as pd


class EstimateFrame(Tk):

    def __init__(self, df):
        super().__init__()

        # Setting the frame
        self.frame = Frame(self, borderwidth=1)
        self.title("sMeta - Разобранная Смета")
        self.center_window()
        self.frame.pack(fill=BOTH, expand=True)

        # Choose window style
        self.style = Style()
        themes = self.style.theme_names()  # all available themes
        self.style.theme_use(themes[1])

        # table
        self.df = df
        self.df = self.get_blocks_as_list()
        self.table = Table(self.frame, dataframe=self.df, showtoolbar=False, showstatusbar=False)
        # resize the columns to fit the data better
        self.table.autoResizeColumns()
        self.table.show()

        # buttons
        self.closeButton = Button(self, text="Закрыть", command=self.destroy)
        self.closeButton.pack(side=RIGHT, padx=5, pady=5)
        self.loadButton = Button(self, text="Сохранить", command=self.save_estimate_to_excel)
        self.loadButton.pack(side=RIGHT)

    #  Save data to excel file
    def save_estimate_to_excel(self):
        Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
        # show an "Save" dialog box and return the path to the file
        filename = dialog.asksaveasfilename(
            defaultextension='.xlsx', filetypes=[("Excel", '*.xlsx')],
            initialdir=os.getcwd(),
            title="Назовите сохраняемый файл")
        self.df = self.table.model.df
        try:
            self.df.to_excel(filename)
        except:
            tkinter.messagebox.showerror("Ошибка", "Ошибка сохранения файла")

    # divide into blocks of estimate

    def get_blocks_as_list(self):
        result = list()
        # find №№
        start = -1
        for i in range(100):
            s = str(self.df.iloc[i, 0])
            isin = s.find('№№', 0, len(s))
            if isin >= 0:
                start = i
                break

        if start == -1:  # start of table not found
            return None

        # create a head
        head = list()
        for i in range(20):
            s = str(self.df.iloc[start, i])
            if s.find('nan', 0, len(s)) > -1:
                break
            head.append(s)

        # start making blocks

        # get indexes of parts
        parts = list()
        for i in range(start, self.df.shape[0]):
            #  search for word Раздел
            s = str(self.df.iloc[i, 0])
            kw = 'Раздел'
            if s.find(kw, 0, len(s)) > -1:
                parts.append(i)

        # form a record
        df_record = pd.DataFrame(columns=head)

        # move through all parts
        for i in parts:
            print(i)
            rec = list()  # one record
            s = str(self.df.iloc[i, 0])
            kw = 'Раздел'
            rec.append(s[len(kw) + 1:])  # all the string without keyword - the name of record

            j = i
            ok = False
            while not ok:
                j += 1
                # analyse line
                s = str(self.df.iloc[j, 0])
                kw = 'Итого'
                if s.find(kw, 0, len(s)) > -1:
                    #  take the number
                    sum = int(self.df.iloc[j, 8])  # in future we should find the right column, now hardcoding
                    sum = str(sum)
                    rec.append(sum)
                    ok = True

                if j > self.df.shape[0]:  # not found
                    break

            print(rec)
            result.append(rec)

        print(result)
        df_record = pd.DataFrame(result)
        print(df_record)

        return df_record







    # try to generate main feature of block to aggregate all others
    def generate_hypotesis(self):
        pass

    # center the frame
    def center_window(self):
        w = 1200
        h = 800
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))