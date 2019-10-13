from tkinter import * 
import os
import openpyxl
from openpyxl import Workbook
from datetime import date

class GUI:
    def __init__(self, master, size = "925x510"):
        self.master = master
        self.master.title("Sign In")
        self.master.geometry(size) #GUI start size

        #set up workbook
        self.wb = openpyxl.load_workbook('data.xlsx')
        self.ws = self.wb.active

        #label
        self.label = Label(master, text="Enter Full Name", font="Helvetica 36 bold")
        self.label.pack()

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
        #setting macros i guess(just makes it easier)
        self.name_col = 2
        self.name_row = self.ws.max_row+1

        #check for if today's column exists, if no make it
        self.currentcol=self.ws.cell(row = 1, column=self.ws.max_column).value
        self.today = str(date.today())
        if (self.currentcol==self.today):
            self.signin_col=self.ws.max_column
        else:
            self.ws.cell(row=1,column=(self.ws.max_column+1)).value = self.today
            self.signin_col=self.ws.max_column
            self.wb.save('data.xlsx')

        #finding name in list and checking them off
        for row in range(2,self.ws.max_row+1):   
            if (self.person.capitalize() == (self.ws.cell(row=row,column=self.name_col).value).capitalize()): #checks if the same
                self.ws.cell(row=row,column=self.signin_col).value = "Present"
                self.wb.save('data.xlsx')




