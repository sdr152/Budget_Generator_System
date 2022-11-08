from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from openpyxl import Workbook
import openpyxl
import os
import datetime as dt
import subprocess

today = dt.datetime.today().date()

root = Tk()
root.title("LISTA DE MATERIALES PARA ELECTRIFICAR")
root.geometry('1000x600')
root.iconphoto(False, PhotoImage(file='peginservice.gif'))
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
        tv2.insert('', 'end', values=detaillst + [1])
        tv1.selection_remove(selected_item)
    else:
        raise "Debe seleccionar un item para agregar al presupuesto."
def remove_fromBudget():
    selected_item = tv2.selection()
    tv2.delete(selected_item)
def generate_Budget():
    # COSTS CALCULATIONS
    iids_for_budget = tv2.get_children()
    detailed_lst = []
    total_item_costs_lst = []
    for iid in iids_for_budget:
        detailed_row = tv2.item(iid)
        values_lst = detailed_row['values']
        detailed_lst.append(values_lst)
        total_cost_per_item = round(1.15*float(values_lst[2])*float(values_lst[3]), 2)
        total_item_costs_lst.append(total_cost_per_item)
    total_costo_materiales = round(sum(total_item_costs_lst), 2)
    mano_de_obra = round(total_costo_materiales * 0.35, 2)
    total_flete = round(total_costo_materiales * 0.10, 2)
    total_imprevistos = round((total_costo_materiales + total_flete) * 0.05, 2)
    TOTAL_PROYECTO = total_costo_materiales + mano_de_obra + total_flete + total_imprevistos
    
    budget_wn = Toplevel(content, borderwidth=20)
    budget_wn.title('Presupuesto')
    budget_wn.iconphoto('False', PhotoImage(file='peginservice.gif'))
    
    canvas = Canvas(budget_wn, width=650, height=600, highlightbackground='red', scrollregion=(0,0, 650, 600))
    save_pdf = ttk.Button(budget_wn, text='Generar pdf')
    sb3 = ttk.Scrollbar(budget_wn, orient=VERTICAL, command=canvas.yview)
    logo_gif1 = PhotoImage(file='peginservice1.gif')
    titles_lst = ['Codigo', 'Material', 'skip', 'Precio unidad', 'Unidades', 'Precio total']
    canvas.create_text(10, 50, text=f'Fecha: {today}', anchor='w',width=100, justify='center', offset='w')
    canvas.create_text(10, 150, text='Codigo', anchor='w', width=100, justify='center', offset='w')
    canvas.create_text(100, 150, text='Material', anchor='w', width=100, justify='center', offset='w')
    canvas.create_text(420, 150, text='Costo unidad', anchor='w', width=100, justify='center', offset='w')
    canvas.create_text(510, 150, text='Cantidad', anchor='w', width=100, justify='center', offset='w')
    canvas.create_text(580, 150, text='Costo total', anchor='w', width=100, justify='center', offset='w')
    canvas.create_line(10, 160, 630, 160, capstyle='projecting')
    for i in range(len(detailed_lst)):
        canvas.create_text(10, 170+i*30, text=detailed_lst[i][0], anchor='w', justify='left', width=70, fill='black')
        canvas.create_text(60, 170+i*30, text=detailed_lst[i][1], anchor='w', justify='left', width=370, fill='black')
        canvas.create_text(450, 170+i*30, text=detailed_lst[i][2], anchor='w', justify='left', width=70, fill='black')
        canvas.create_text(530, 170+i*30, text=detailed_lst[i][3], anchor='w', justify='left', width=70, fill='black')
        canvas.create_text(590, 170+i*30, text=total_item_costs_lst[i], anchor='w', justify='left', width=70, fill='red')
    
    print(canvas.winfo_screenwidth())
    print(canvas.winfo_width(), canvas.winfo_height())
    print(logo_gif.width(), logo_gif.height())
    canvas.grid(column=0, row=0, columnspan=6, padx=5, pady=5)
    sb3.grid(column=7, row=0, rowspan=2, sticky='ns')
    save_pdf.grid(column=0, row=1)
    budget_wn.config(yscrollcomman=sb3.set)
    canvas.create_image(540, 70, image=logo_gif)
    canvas.update()
    canvas.postscript(file='tmp.ps', colormode='color')
    process = subprocess.Popen(["ps2pdf", "tmp.ps", "new_pdf.pdf"], shell=True)
    process.wait()
    os.remove("tmp.ps")
    
def gen_pdf():
    pass
def on_closing():
    if messagebox.askokcancel('Quit', 'Do you wanto to quit?'):
        root.destroy()
def fill_treeview():
    for row in ws.values:
        id = tv1.insert('', 'end', values=[row[0], row[1], row[2]])
def update_num_units(event):
    region_clicked = tv2.identify_region(event.x, event.y)
    identified_col = tv2.identify_column(event.x)
    identified_row = tv2.identify_row(event.y)
    if region_clicked in ('cell') and identified_col == '#4':
        selected_id = tv2.selection()
        selected_row = tv2.item(selected_id)
        detaillst = selected_row['values']
        bbox = tv2.bbox(selected_id, 'Unidades')
        units_entry = ttk.Entry(tv2, width=bbox[2])
        units_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        units_entry.insert(0, detaillst[3])
        units_entry.select_range(0, END)
        units_entry.focus()
        units_entry.bind("<FocusOut>", on_focus_out)
        units_entry.bind("<Return>", on_return)       
def on_focus_out(event):
    event.widget.destroy()
def on_return(event):
    vl = event.widget.get()
    selected_id = tv2.selection()
    selected_id_index = tv2.index(selected_id)
    selected_row = tv2.item(selected_id)
    detaillst = selected_row['values']
    detaillst[3] = vl
    tv2.delete(selected_id)
    tv2.insert('', index=selected_id_index, iid=selected_id, values=detaillst)
    event.widget.destroy()
    
# CREATE WIDGETS
logo_gif = PhotoImage(file='peginservice.gif')
logo_lb = ttk.Label(content, image=logo_gif, relief='ridge') #relief: flat, groove, raised, ridge, solid, or sunken
titlelbl = ttk.Label(content, text='LISTA DE MATERIALES PARA ELECTRIFICAR', justify='center') 
codelbl = ttk.Label(content, text='Codigo:')
matlbl = ttk.Label(content, text='Material:')
pricelbl = ttk.Label(content, text='Precio unidad:')

codeEntry = ttk.Entry(content, textvariable=code)
matEntry = ttk.Entry(content, textvariable=mat)
priceEntry = ttk.Entry(content, textvariable=price)

Add = ttk.Button(content, text='Agregar', command=add_toDb, width=25)
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
logo_lb.grid(column=5, row=1, rowspan=3, padx=5, pady=5, sticky='we')

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

tv2.bind("<Double-1>", update_num_units)

root.protocol('WM_DELETE_WINDOW', on_closing)
root.mainloop()