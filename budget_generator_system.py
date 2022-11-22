from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from openpyxl import Workbook
import openpyxl
import os
import datetime as dt
import subprocess
import time
import shutil
from PIL import Image

root = Tk()
root.title("LISTA DE MATERIALES PARA ELECTRIFICAR")
root.geometry('1000x600')
root.iconphoto(False, PhotoImage(file='peginservice.gif'))
# CREATE A MAIN FRAME
content = ttk.Frame(root, padding=(5,5,12,12), borderwidth=5)
content.grid(column=0, row=0, sticky='nsew')

# VARIABLES
page_cap = 1000
today = dt.datetime.today().date()
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
    for id, rw in enumerate(ws.values):
        if detaillst[0]==rw[0]:
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
def update_tv1_entry(event):
    selected_item = event.widget.selection()
    #selected_item = tv1.selection()
    if selected_item:
        #detaillst = tv1.item(selected_item)['values']
        detaillst = event.widget.item(selected_item)['values']
        code.set(detaillst[0])
        mat.set(detaillst[1])
        price.set(detaillst[2])
def generate_Budget():
    def gen_pdf(*args):
        # Convert to PDF
        current_dir = os.getcwd()
        #current_dir = current_dir.replace("\\" , '/')
        print(current_dir)
        os.makedirs(f'C:/Users/Samuel Ramos/Documents/{cl_name.get()}', exist_ok=True)
        #os.chdir(f'C:/Users/Samuel Ramos/Documents/{cl_name.get()}')
        
        for i, cnv in enumerate(canvas_lst):
            cnv.create_image(630, 40, image=logo_gif)
            cnv.update()
            cnv.postscript(file='tmp.ps', fontmap='-*-Courier-Bold-R-Normal--*-120-*', colormode='color', pagex=300, pagey=420, height=1000, width=700)
            process = subprocess.Popen(["ps2pdf", "tmp.ps", f"Budget_{i}.pdf"], shell=True)
            process.wait()
            os.remove('tmp.ps')
            current_dir = current_dir.replace('\\', '/')
            #os.rename(current_dir+"/"+f"Budget_{i}.pdf", f'C:/Users/Samuel Ramos/Documents/{cl_name.get()}/Budget_{i}.pdf')
            os.replace(current_dir+"/"+f"Budget_{i}.pdf", f'C:/Users/Samuel Ramos/Documents/{cl_name.get()}/Budget_{i}.pdf')
            #shutil.move(current_dir+"/"+f"Budget_{i}.pdf", f'C:/Users/Samuel Ramos/Documents/{cl_name.get()}/Budget_{i}.pdf')
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
    TOTAL_PROYECTO = round((total_costo_materiales + mano_de_obra + total_flete + total_imprevistos), 2)
    resumen_costos = [total_costo_materiales, mano_de_obra, total_flete, total_imprevistos, TOTAL_PROYECTO]
    
    budget_wn = Toplevel(content, borderwidth=20, width=700)
    budget_wn.title('Presupuesto')
    budget_wn.iconphoto('False', PhotoImage(file='peginservice.gif'))

    logo_gif = PhotoImage(file='peginservice.gif', palette=1)

    # Create a main frame
    main_frame = Frame(budget_wn, width=700) #h 500
    main_frame.pack(fill=BOTH, expand=1)
    
    # Create a save pdf button
    save_pdf = ttk.Button(main_frame, text='Guardar como PDF', command=gen_pdf).pack(side=BOTTOM, fill=X)
    
    # Create a canvas
    canvas = Canvas(main_frame, highlightbackground='black', bg='gray', width=700) 
    canvas.pack(side=LEFT, fill=BOTH, expand=1)
    
    # Create a Scrollbar
    sb3 = ttk.Scrollbar(main_frame, orient=VERTICAL, command=canvas.yview)
    sb3.pack(side=RIGHT, fill=Y)
 
    # Configure canvas
    canvas.configure(yscrollcommand=sb3.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

    # Add another frame inside canvas
    second_frame = Frame(canvas, height=800) # h 800
    
    # Add new frame to a window in the canvas
    num_pages = len(detailed_lst)//20 + 1
    canvas.create_window((0,0), window=second_frame, anchor='nw', height=page_cap*num_pages) #h 800
    
    costs_labels = ['Costo total por materiales:', 'Total mano de obra:', 'Total Flete:','Costo por imprevistos:', 'COSTO TOTAL DEL PROYECTO:']
    header_labels = ['Fecha:', 'Nombre de cliente:', 'R.T.N.:', 'No. Factura:', 'No. Pagina:']
    heading_labels = [('Codigo', 10), ('Material', 100), ('Costo Unidad Lps.', 430), ('Cantidad', 540), ('Costo Total Lps., ISV incluido', 600)]
    
    def on_enter1(event):
        vl = event.widget.get()
        for canvas in canvas_lst:
            canvas.create_text(130, 50, text=vl, anchor='w', width=270, justify='left')
        event.widget.destroy()
    def on_enter2(event):
        vl = event.widget.get()
        for canvas in canvas_lst:
            canvas.create_text(130, 70, text=vl, anchor='w', width=270, justify='left')
        event.widget.destroy()
    def on_enter3(event):
        vl = event.widget.get()
        for canvas in canvas_lst:
            canvas.create_text(130, 90, text=vl, anchor='w', width=270, justify='left')
        event.widget.destroy()
    
    sublsts = list(create_sublists(detailed_lst, 20))
    total_item_costs_sublsts = list(create_sublists(total_item_costs_lst, 20))
    canvas_lst = list(create_canvas(sublsts, second_frame))
    
    # Fecha
    canvas.create_text(150, 30, text=today, anchor='w', width=270, justify='left')
    # Client name entry
    cl_name = StringVar()
    client_name_entry = ttk.Entry(canvas_lst[0], textvariable=cl_name)
    client_name_entry.place(x=115, y=40, width=250, height=20)
    client_name_entry.bind("<Return>", on_enter1)
    
    # RTN entry
    rtn = StringVar()
    rtn_entry = ttk.Entry(canvas_lst[0], textvariable=rtn)
    rtn_entry.place(x=115, y=60, width=250, height=20)
    rtn_entry.bind('<Return>', on_enter2)
    # N. Factura entry
    factura = StringVar()
    factura_entry = ttk.Entry(canvas_lst[0], textvariable=factura)
    factura_entry.place(x=115, y=80, width=250, height=20)
    factura_entry.bind('<Return>', on_enter3)
    
    for idx in range(len(canvas_lst)):
        sub_lst = sublsts[idx]
        cost_sub_lst = total_item_costs_sublsts[idx]
        cv = canvas_lst[idx]
        cv.pack(side=TOP, fill=BOTH, expand=1)
        
        # Put the date on each canvas
        cv.create_text(130, 30, text=today, anchor='w', width=270, justify='left')

        # Put the page number on each canvas
        cv.create_text(130, 110, text=f'{idx+1} of {len(canvas_lst)}', anchor='w', width=270, justify='left')
        
        # Put the page header on each canvas
        for i in range(len(header_labels)):
            cv.create_text(10, i*20+30, text=header_labels[i], anchor='w', width=300, justify='left')
        
        # Put company information on canvas
        cv.create_text(690, 85, text='Cel: +504 9799-2662', anchor='e', justify='right')
        cv.create_text(690, 100, text='Email: ventas@peginservice.com', anchor='e', justify='right')
        cv.create_text(690, 115, text='Col. Miraflores, Calle Guanaja # 1884, Tegucigalpa, Honduras', anchor='e', justify='right')
        # Put the table header on each canvas
        for i in range(len(heading_labels)):
            cv.create_text(heading_labels[i][1], 155, text=heading_labels[i][0], anchor='w', width=100, justify='center')
        cv.create_line(10, 170, 690, 170, capstyle='round')
        
        # Fill out the table for each canvas
        for i in range(len(sub_lst)):
            cv.create_text(10, 180+i*30, text=sub_lst[i][0], anchor='w', justify='left', width=70, fill='black')
            cv.create_text(60, 180+i*30, text=sub_lst[i][1], anchor='w', justify='left', width=370, fill='black')
            cv.create_text(460, 180+i*30, text=sub_lst[i][2], anchor='w', justify='left', width=70, fill='black')
            cv.create_text(560, 180+i*30, text=sub_lst[i][3], anchor='w', justify='left', width=70, fill='black')
        for i in range(len(cost_sub_lst)):
            cv.create_text(630, 180+i*30, text=cost_sub_lst[i], anchor='w', justify='left', width=70, fill='red')
        if idx == len(canvas_lst)-1:
            cv.create_text(10, 180+len(sub_lst)*30, text="Resumen de Costos: ", anchor='w', justify='left', width=150, fill='black')
            line = cv.create_line(10, 190+len(sub_lst)*30, 690, 190+len(sub_lst)*30, capstyle='butt')
            for i in range(len(costs_labels)):
                cv.create_text(530, 210+len(sub_lst)*30+i*20, text=costs_labels[i], anchor='e', justify='right', width=300, fill='black')
                cv.create_text(570, 210+len(sub_lst)*30+i*20, text=f'Lps. {resumen_costos[i]}', anchor='w', justify='left', width=70, fill='red')
                if i>3:
                    cv.create_line(20, 310+len(sub_lst)*30+i*20, 150, 310+len(sub_lst)*30+i*20)
                    cv.create_line(200, 310+len(sub_lst)*30+i*20, 330, 310+len(sub_lst)*30+i*20)
                    cv.create_text(20, 320+len(sub_lst)*30+i*20, text="PeginService", anchor='w', justify='left', width=70, fill='red')
                    cv.create_text(200, 320+len(sub_lst)*30+i*20, text="Cliente", anchor='w', justify='left', width=70, fill='red')
def create_canvas(lst, frame):
    for idx in range(len(lst)):
        canvas = Canvas(frame, highlightbackground='red', bg='yellow', width=700, height=page_cap)
        yield canvas
    
def create_sublists(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]
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
        units_entry.bind("<Double-Button-3>", on_return)     
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
def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    
    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    # reverse sort next time
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

# CREATE WIDGETS
# FIRST WINDOW WIDGETS
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

cols1 = ['Código', 'Material', 'Costo unidad'] 
cols = ['Código', 'Material', 'Costo unidad', 'Unidades']

tv1 = ttk.Treeview(content, columns=cols[:3], show='headings', height=15)
tv1.column(cols[0], width=15)
tv1.column(cols[2], width=30)
for col in cols[:3]:
    tv1.heading(col, text=col, command=lambda: treeview_sort_column(tv1, 'Material', False))
sb1 = ttk.Scrollbar(content, orient=VERTICAL, command=tv1.yview)
tv1.config(yscrollcommand=sb1.set)
tv1.bind("<Double-1>", update_tv1_entry)

tv2 = ttk.Treeview(content, columns=cols, show='headings', height=15)
tv2.column(cols[0], width=70)
tv2.column(cols[1], width=150)
tv2.column(cols[2], width=90)
tv2.column(cols[3], width=80)
for col2 in tv2['column']:
    tv2.heading(col2, text=col2, command=lambda: treeview_sort_column(tv2, 'Material', False))
sb2 = ttk.Scrollbar(content, orient=VERTICAL, command=tv2.yview)
tv2.config(yscrollcommand=sb2.set)
# SECOND WINDOW WIDGETS

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
logo_lb.grid(column=5, row=1, rowspan=3, padx=5, pady=5, sticky='e')

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