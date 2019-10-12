from tkinter import * 
import os

class GUI:
    def __init__(self, master, size = "925x510"):
        self.master = master
        self.master.title("Sign In")
        self.master.geometry(size) #GUI start size

        #input box
        self.inputbox = Entry(self.master, font = "Helvetica 44 bold", width=25)
        self.inputbox.pack()

        #signin button
        self.signin_button = Button(self.master, font = "Helvetica 36 bold", text="Sign In", command=self.signin)
        self.signin_button.pack()

    #functions

    #signin
    def signin(self):
        identity = self.inputbox.get()
        self.inputbox.delete(0, 'end')
