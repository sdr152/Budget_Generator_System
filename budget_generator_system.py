from distutils.log import error
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from openpyxl import Workbook
import openpyxl
import pandas as pd
import os

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
    if code.get() != '' and mat.get() != '' and price.get() != '':
        id = tv1.insert('', 'end', values=[code.get(), mat.get(), price.get()])
        ws.append([code.get(), mat.get(), price.get()])
        wb.save('database.xlsx')
        code.set(''), mat.set(''), price.set('')
def remove_fromDb():
    selected_item = tv1.selection()
    detaillst = tv1.item(selected_item)['values']
    print(detaillst)
    for id, rw in enumerate(ws.values):
        print(rw)
        if detaillst[0]==rw[0]:
            print('SAME :)')
            ws.delete_rows(id+1)
            break
    wb.save('database.xlsx')
    tv1.delete(selected_item)
def add_toBudget():
    selected_item = tv1.selection()
    if selected_item:
        detaillst = tv1.item(selected_item)['values']
        tv2.insert('', 'end', values=detaillst)
        tv1.selection_remove(selected_item)
    else:
        raise "Debe seleccionar un item para agregar al presupuesto."
def remove_fromBudget():
    selected_item = tv2.selection()
    tv2.delete(selected_item)
def generate_Budget():
    budget_wn = Toplevel(content)
    budget_wn.title('Presupuesto')
    budget_wn.config(height=600, width=500)
def on_closing():
    if messagebox.askokcancel('Quit', 'Do you wanto to quit?'):
        root.destroy()
def fill_treeview():
    for row in ws.values:
        id = tv1.insert('', 'end', values=[row[0], row[1], row[2]])
        

# CREATE WIDGETS
titlelbl = ttk.Label(content, text='LISTA DE MATERIALES PARA ELECTRIFICAR', justify='center')
codelbl = ttk.Label(content, text='Codigo:')
matlbl = ttk.Label(content, text='Material:')
pricelbl = ttk.Label(content, text='Precio unidad:')

codeEntry = ttk.Entry(content, textvariable=code)
matEntry = ttk.Entry(content, textvariable=mat)
priceEntry = ttk.Entry(content, textvariable=price)

Add = ttk.Button(content, text='Añadir', command=add_toDb, width=25)
Remove = ttk.Button(content, text='Eliminar', command=remove_fromDb, width=25)
AddtoBudget = ttk.Button(content, text='Agregar a Presupuesto', command=add_toBudget, width=25)
RemovefromBudget = ttk.Button(content, text='Eliminar de Presupuesto', command=remove_fromBudget, width=25)
Generate = ttk.Button(content, text='Generar', command=generate_Budget, width=25)

cols = ['Código', 'Material', 'Costo unidad', 'Unidades']
tv1 = ttk.Treeview(content, columns=cols[:3], show='headings', height=15)
tv1.column(cols[0], width=15)
tv1.column(cols[2], width=30)
for col in tv1['column']:
    tv1.heading(col, text=col)
sb1 = ttk.Scrollbar(content, orient=VERTICAL, command=tv1.yview)
tv1.config(yscrollcommand=sb1.set)

tv2 = ttk.Treeview(content, columns=cols, show='headings', height=15)
tv2.column(cols[0], width=70)
tv2.column(cols[1], width=150)
tv2.column(cols[2], width=90)
tv2.column(cols[3], width=80)
for col in tv2['column']:
    tv2.heading(col, text=col)
sb2 = ttk.Scrollbar(content, orient=VERTICAL, command=tv2.yview)
tv2.config(yscrollcomman=sb2.set)

# GRID WIDGETS
titlelbl.grid(column=0, row=0, columnspan=8, padx=5, pady=5, sticky=N)
codelbl.grid(column=0, row=1, padx=5, pady=5, sticky=W)
matlbl.grid(column=0, row=2, padx=5, pady=5, sticky=W)
pricelbl.grid(column=0, row=3, padx=5, pady=5, sticky=W)
codeEntry.grid(column=1, row=1, padx=5, pady=5,  sticky='we')
matEntry.grid(column=1, row=2, columnspan=2, padx=5, pady=5, sticky='we')
priceEntry.grid(column=1, row=3, padx=5, pady=5, sticky='we')
Add.grid(column=0, row=4, padx=5, pady=5, sticky='we')
Remove.grid(column=1, row=4, padx=5, pady=5, sticky='we')
AddtoBudget.grid(column=2, row=4, padx=5, pady=5, sticky='we')
RemovefromBudget.grid(column=4, row=4, padx=5, pady=5, sticky='we')
Generate.grid(column=5, row=4, padx=5, pady=5, sticky='we')
tv1.grid(column=0, row=5, columnspan=3, padx=5, pady=5, sticky="nsew")
tv2.grid(column=4, row=5, columnspan=3, padx=5, pady=5, sticky='nsew')
sb1.grid(column=3, row=5, padx=5, pady=5, sticky="ns")
sb2.grid(column=7, row=5, padx=5, pady=5, sticky='ns')

root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(0, weight=1)

# CHECK IF THERE IS A DB PRESENT
path = 'database.xlsx'
isExist = os.path.exists(path)
if isExist:
    wb = openpyxl.load_workbook('database.xlsx')
    ws = wb.active
    fill_treeview()
    
else:
    wb = Workbook()
    ws = wb.active
    
#wb.save('database.xlsx')



#for row in DB.iter_rows(max_col=4, max_row=6):
#    for v in row:
#        print(v.value)
#rowsit = DB.iter_rows()
#print(len(list(rowsit)))
#df = pd.DataFrame([('cod1', 'ramses', 99)], columns=['Código', 'Material', 'Costo unidad'])
#DB = DB.append(df)
#DB.loc[-1] = ['qqq', 'rejas', 43, 999, 9]
#print(DB)
#DB.to_csv('database.csv')
root.protocol('WM_DELETE_WINDOW', on_closing)
root.mainloop()