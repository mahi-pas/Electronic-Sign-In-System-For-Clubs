from tkinter import * 
import os
import pandas as pd
import openpyxl
from openpyxl import Workbook

class GUI:
    def __init__(self, master, size = "925x510"):
        self.master = master
        self.master.title("Sign In")
        self.master.geometry(size) #GUI start size

        #set up workbook
        self.wb = openpyxl.load_workbook('data.xlsx')
        self.ws = self.wb.active

        #input box
        self.inputbox = Entry(self.master, font = "Helvetica 44 bold", width=25)
        self.inputbox.pack()

        #signin button
        self.signin_button = Button(self.master, font = "Helvetica 36 bold", text="Sign In", command=self.signin)
        self.signin_button.pack()

    #functions

    #signin
    def signin(self):
        self.person = self.inputbox.get()
        self.inputbox.delete(0, 'end')

        #excel stuff
        self.name_col = 2
        self.name_row = self.ws.cell(row=1,column=1).value
        print(self.name_row)
