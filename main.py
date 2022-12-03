import pandas as pd
import convert_numbers as cf
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

class Program (object):

    def __init__(self):
        self.INFILEPATH = ''
        self.OUTFILEPATH = ''
        self.CHOICES = ['en', 'ar']
        self.window = tk.Tk()
        self.window.title('Numbers Converter')
        self.window.resizable(width=False, height=False)
        self.tolang = tk.StringVar(self.window)
        self.tolang.set('ar')
        self.fromlang = tk.StringVar(self.window)
        self.fromlang.set('en')

        self.direcButton = tk.Button(
            text = 'Browse File',
            width = 25,
            height = 1,
            command = self.__askdirec
        )

        self.direcLabel = tk.Label(
            text = "Input Location: " + self.INFILEPATH
        )

        self.convButton = tk.Button(
            text = 'Covert Now',
            width = 25,
            height = 1,
            command = self.__convertnumbers
        )

        self.fromOptions = tk.OptionMenu(self.window,
                                         self.fromlang,
                                         *self.CHOICES
                                         )

        self.toOptions = tk.OptionMenu(self.window,
                                         self.tolang,
                                         *self.CHOICES
                                         )

        self.direcLabel1 = tk.Label(
            text = "Output Location: " + self.OUTFILEPATH
        )
        self.__start()

    def __start(self):
        self.direcButton.pack()
        self.fromOptions.pack()
        self.direcLabel.pack()
        self.convButton.pack()
        self.toOptions.pack()
        self.direcLabel1.pack()
        self.window.mainloop()

    def __askdirec(self) -> str:
        self.OUTFILEPATH = ''
        self.INFILEPATH = filedialog.askopenfilename(
            parent = self.window,
            initialdir = os.getcwd(),
            title = 'Please select a directory',
            filetypes=(
                ("Excel files", "*.xlsx"),
                ("all files", "*.*")
                        )
                                        )
        self.direcLabel.config(text = "Input Location: " + self.INFILEPATH)
        temp = self.INFILEPATH.split('/')
        temp[len(temp) - 1] = 'New_Numbers.xlsx'
        for i,x in enumerate(temp):
            self.OUTFILEPATH += x + '/' if i != len(temp) - 1 else x
        self.direcLabel1.config(text = "Output Location: " + self.OUTFILEPATH)
        return self.INFILEPATH

    def __convertnumbers(self):
        try:
            df = pd.read_excel(self.INFILEPATH)

            if self.tolang.get() == self.fromlang.get():
                raise Exception("ERROR")

            elif self.fromlang.get() == 'en' and self.tolang.get() == 'ar':
                df["Numbers"] = df["Numbers"].astype(str)
                for i, x in enumerate(df["Numbers"]):
                    df["Numbers"][i] = cf.english_to_arabic(x)
                df.rename(columns={"Numbers": "الارقام"}, inplace=True)

            elif self.fromlang.get() == 'ar' and self.tolang.get() == 'en':
                df["الارقام"] = df["الارقام"].astype(str)
                for i, x in enumerate(df["الارقام"]):
                    df["الارقام"][i] = cf.arabic_to_english(x)
                df.rename(columns={"الارقام": "Numbers"}, inplace=True)

            df.to_excel(self.OUTFILEPATH)

            messagebox.showinfo("DONE!", "DONE!!")

            self.INFILEPATH = ''
            self.OUTFILEPATH = ''

            self.direcLabel.config(text="Input Location: " + self.INFILEPATH)
            self.direcLabel1.config(text="Output Location: " + self.OUTFILEPATH)

        except:
            self.direcLabel.config(text="Input Location: " + "ERROR")
            self.direcLabel1.config(text="Output Location: " + "ERROR")


Program()
