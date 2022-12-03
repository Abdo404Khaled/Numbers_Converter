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
        self.window = tk.Tk()
        self.window.title('Numbers Converter')
        self.window.resizable(width=False, height=False)

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

        self.direcLabel1 = tk.Label(
            text = "Output Location: " + self.OUTFILEPATH
        )
        self.__start()

    def __start(self):
        self.direcButton.pack()
        self.direcLabel.pack()
        self.convButton.pack()
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
            df["Numbers"] = df["Numbers"].astype(str)

            for i, x in enumerate(df["Numbers"]):
                df["Numbers"][i] = cf.english_to_arabic(x)

            df.rename(columns={"Numbers": "الارقام"}, inplace=True)
            df.to_excel(self.OUTFILEPATH)
            messagebox.showinfo("DONE!", "DONE!!")
            self.INFILEPATH = ''
            self.OUTFILEPATH = ''
            self.direcLabel.config(text="Input Location: " + self.INFILEPATH)
            self.direcLabel1.config(text="Output Location: " + self.OUTFILEPATH)
        except:
            self.direcLabel.config(text="Input Location: " + "PLEASE CHOOSE LOCATION")
            self.direcLabel1.config(text="Output Location: " + "ERROR")


Program()
