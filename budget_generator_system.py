from tkinter import *
from tkinter import ttk
from csv
import pandas as pd

root = Tk()
root.title("LISTA DE MATERIALES PARA ELECTRIFICAR")
root.geometry('1000x600')

# CREATE A MAIN FRAME
content = ttk.Frame(root, padding=(5,5,12,12), borderwidth=5)
content.grid(column=0, row=0, sticky='nsew')

# VARIABLES
code = StringVar()
mat = StringVar()
price = StringVar()

# FUNCTIONS
def add_toDb():
    pass
def remove_fromDb():
    pass
def add_toBudget():
    pass
def remove_fromBudget():
    pass
def generate_Budget():
    pass



root.mainloop()