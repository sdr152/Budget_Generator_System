from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkinter.font import Font
from openpyxl import Workbook
from PyPDF2 import PdfFileMerger
import openpyxl
import os
import datetime as dt
import subprocess
import win32api

root = Tk()
root.title("GENERADOR DE COTIZACIONES")
root.geometry('1150x600')
root.iconphoto(False, PhotoImage(file='images/peginservice.gif'))
# CREATE A MAIN FRAME
content = ttk.Frame(root, padding=(5,5,12,12), borderwidth=5)
content.grid(column=0, row=0, sticky='nsew')

# FUNCTIONS
def add_toDb():
    try:
        if code.get() != '' and mat.get() != '' and price.get() != '' and unit.get() != '':
            id = tv1.insert('', 'end', values=[code.get(), mat.get(), unit.get(), float(price.get())])
            ws.append([code.get(), mat.get(), unit.get(), price.get()])
            wb.save(f'{path_to_usr}/Cotizaciones PegIn/database.xlsx')
            code.set(''), mat.set(''), unit.set(''), price.set('')
    except ValueError:
        print('Por favor, ingrese dato numerico en la casilla de precio.')
def remove_fromDb():
    selected_item = tv1.selection()
    if selected_item:    
        detaillst = tv1.item(selected_item)['values']
        for id, rw in enumerate(ws.values):
            if detaillst[0]==rw[0]:
                ws.delete_rows(id+1)
                break
        wb.save(f'{path_to_usr}/Cotizaciones PegIn/database.xlsx')
        tv1.delete(selected_item)
def add_toBudget():
    selected_item = tv1.selection()
    if selected_item:
        detaillst = tv1.item(selected_item)['values']
        tv2.insert('', 'end', values=detaillst + [1])
        tv1.selection_remove(selected_item)
def remove_fromBudget():
    selected_item = tv2.selection()
    if selected_item:
        tv2.delete(selected_item)
def update_tv1_entry(event):
    selected_item = event.widget.selection()
    if selected_item:
        detaillst = event.widget.item(selected_item)['values']
        code.set(detaillst[0])
        mat.set(detaillst[1])
        unit.set(detaillst[2])
        price.set(detaillst[3])
def create_canvas(lst, frame):
    for idx in range(len(lst)):
        canvas = Canvas(frame, highlightbackground='black', bg='#fdffc7', width=800, height=page_cap)
        yield canvas
def create_sublists(lst, size):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]
def on_closing():
    if messagebox.askokcancel('Quit', 'Do you wanto to quit?'):
        root.destroy()
def fill_treeview():
    for row in ws.values:
        id = tv1.insert('', 'end', values=[row[0], row[1], row[2], float(row[3])])
def update_num_units(event):
    region_clicked = tv2.identify_region(event.x, event.y)
    identified_col = tv2.identify_column(event.x)
    if region_clicked in ('cell') and identified_col == '#5':
        selected_id = tv2.selection()
        selected_row = tv2.item(selected_id)
        detaillst = selected_row['values']
        bbox = tv2.bbox(selected_id, 'Cantidad')
        units_entry = ttk.Entry(tv2, width=bbox[2])
        units_entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        units_entry.insert(0, detaillst[4])
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
    detaillst[4] = vl
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
def generate_Budget():
    def pago_lbl():
        if not canvas_lst[0].itemcget('tagpago', 'text'):
            canvas_lst[0].create_text(400, 20, text=pago.get(), anchor='w', width=270, justify='left', font=my_font2, tags='tagpago')
        else:
            canvas_lst[0].delete('tagpago')
            canvas_lst[0].create_text(400, 20, text=pago.get(), anchor='w', width=270, justify='left', font=my_font2, tags='tagpago')
    def on_enter1(event):
        vl = event.widget.get()
        for canvas in canvas_lst:
            canvas.create_text(95, 50, text=vl, anchor='w', width=270, justify='left', font=my_font2)
        event.widget.destroy()
    def on_enter2(event):
        vl = event.widget.get()
        for canvas in canvas_lst:
            canvas.create_text(95, 70, text=vl, anchor='w', width=270, justify='left', font=my_font2)
        event.widget.destroy()
    def save_as():
        # Convert to PDF
        if cl_name.get() == '':
            print("Client must have a name!")
            return

        print("STEP 1")
        # 1. Check if directory exists or create a directory to save pdf inside of.
        if not os.path.exists(f'{path_to_usr}/Cotizaciones PegIn/Cotizaciones - {dt.datetime.today().strftime("%B")}'):
            os.makedirs(f'{path_to_usr}/Cotizaciones PegIn/Cotizaciones - {dt.datetime.today().strftime("%B")}', exist_ok=True)
        
        filename_to_save = filedialog.asksaveasfilename(
            initialdir=f'{path_to_usr}/Cotizaciones PegIn/Cotizaciones - {dt.datetime.today().strftime("%B")}', title='Save as',
            defaultextension='.pdf', initialfile=f'{cl_name.get()}_cotizacion'
        )
        
        print("STEP 2: Create and move pdf files")
        # 2. Convert canvas to pdf and move pdfs to directory created in #1.
        os.chdir(os.path.dirname(filename_to_save))
        for i, cnv in enumerate(canvas_lst):
            cnv.create_image(730, 40, image=logo_gif)
            cnv.update()
            cnv.postscript(file='tmp.ps', fontmap='-*-Courier-Bold-R-Normal--*-120-*', colormode='color', pagex=300, pagey=420, height=1000, width=800)
            process = subprocess.Popen(["ps2pdf", "tmp.ps", f"pagina_{i}.pdf"], shell=True)
            process.wait()
            os.remove('tmp.ps')
            
        print('STEP 3: Merge files ')        
        merger = PdfFileMerger()
        path_to_files = os.path.dirname(filename_to_save) + '/'
        for root, dirs, file_names in os.walk(path_to_files):        
                for file_name in file_names:
                    if file_name.startswith("pagina"):
                        merger.append(path_to_files + file_name)    
        merger.write(filename_to_save)
        merger.close()
        
        print("STEP 4: Clean paginas")
        # 4. Delete remaining pdf files.
        for root, dirs, file_names in os.walk(path_to_files):
            for file_name in file_names:
                if file_name.startswith("pagina"):
                    os.remove(path_to_files + file_name)
        print("FINISH CLEANING")

        # 5. Append path of pdf file to a list.
        files_to_print_lst.append(filename_to_save)

    def print_hard_copy():
        if files_to_print_lst:
            path_to_file = files_to_print_lst[-1]
            win32api.ShellExecute(0, 'print', path_to_file, None, '.', 0)
            print(files_to_print_lst)
        else:
            messagebox.showerror("Error!", "Guarde una copia de su documento pdf antes de imprimir.")
    
    hr = dt.datetime.today().hour
    mn = dt.datetime.today().minute
    sc = dt.datetime.today().second
    files_to_print_lst = []
    
    # COSTS CALCULATIONS
    if not tv2.get_children():
        return
    iids_for_budget = tv2.get_children()
    detailed_lst = []
    total_item_costs_lst = []
    isv = 0.0
    for iid in iids_for_budget:
        detailed_row = tv2.item(iid)
        values_lst = detailed_row['values']
        detailed_lst.append(values_lst)
        total_cost_per_item = round(float(values_lst[3]) * float(values_lst[4]), 2)
        total_item_costs_lst.append(total_cost_per_item)
        isv += round(0.15 * float(values_lst[3]) * float(values_lst[4]), 2)
    isv = round(isv, 2)
    total_bruto = round(sum(total_item_costs_lst), 2)
    total_discount = round(-float(discount.get()) / 100 * total_bruto, 2)
    total_neto = round(total_bruto + isv + total_discount, 2)
    resumen_costos = [total_bruto, isv, total_discount, total_neto]
    
    budget_wn = Toplevel(content, borderwidth=20, width=800, height=1000)
    budget_wn.title('Cotizacion')
    budget_wn.iconphoto('False', PhotoImage(file=f'{path_to_program}/images/peginservice.gif'))
    # disable parent window while child window is open
    budget_wn.grab_set()
    logo_gif = PhotoImage(file=f'{path_to_program}/images/peginservice.gif', palette=1)
    
    # Create a main frame
    main_frame = Frame(budget_wn, width=800, height=1000)
    main_frame.pack(fill=BOTH, expand=1)
    
    # Create a save pdf button
    pago = StringVar()
    forma_pago = ttk.Checkbutton(main_frame, text='Pago a credito', variable=pago, onvalue='*Pago a credito*', offvalue='*Pago por contado*', command=pago_lbl)
    forma_pago.grid(column=2, row=0, sticky='w')
    
    print_hc = ttk.Button(main_frame, text='Imprimir', command=print_hard_copy)
    print_hc.grid(column=1, row=0, sticky='w')
    save_as = ttk.Button(main_frame, text='Save as', command=save_as)
    save_as.grid(column=0, row=0, sticky='w')
    # Create a canvas
    canvas = Canvas(main_frame, highlightbackground='black', bg='gray', width=800, height=500) 
    canvas.grid(column=0, row=1, columnspan=5)
    # Create a Scrollbar
    sb3 = ttk.Scrollbar(main_frame, orient=VERTICAL, command=canvas.yview)
    sb3.grid(column=5, row=1, sticky='ns')
 
    # Configure canvas
    canvas.configure(yscrollcommand=sb3.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

    # Add another frame inside canvas
    second_frame = Frame(canvas, height=800) 
    
    # Add new frame to a window in the canvas
    num_pages = len(detailed_lst)//20 + 1
    canvas.create_window((0,0), window=second_frame, anchor='nw', height=page_cap*num_pages) 
    
    costs_labels = ['Costo total bruto:','Total impuesto sobre venta:' ,'Descuento costo total:', 'COSTO TOTAL NETO:']
    header_labels = ['Fecha:', 'Cliente:', 'R.T.N.:', 'No. Cotizacion:', 'No. Pagina:']
    heading_labels = [('Codigo', 10), ('Material', 100), ('Unidad', 490),('Costo Unidad Lps.', 540), ('Cantidad', 640), ('Costo Total Lps., ISV incluido', 700)]
    
    sublsts = list(create_sublists(detailed_lst, 20))
    total_item_costs_sublsts = list(create_sublists(total_item_costs_lst, 20))
    canvas_lst = list(create_canvas(sublsts, second_frame))
    
    # Client name entry
    cl_name = StringVar()
    client_name_entry = ttk.Entry(canvas_lst[0], textvariable=cl_name)
    client_name_entry.place(x=85, y=40, width=250, height=20)
    client_name_entry.bind("<Return>", on_enter1)
    
    # RTN entry
    rtn = StringVar()
    rtn_entry = ttk.Entry(canvas_lst[0], textvariable=rtn)
    rtn_entry.place(x=85, y=60, width=250, height=20)
    rtn_entry.bind('<Return>', on_enter2)
    
    for idx in range(len(canvas_lst)):
        sub_lst = sublsts[idx]
        cost_sub_lst = total_item_costs_sublsts[idx]
        cv = canvas_lst[idx]
        cv.pack(side=TOP, fill=BOTH, expand=1)
        
        # Put the date on each canvas
        cv.create_text(95, 30, text=today, anchor='w', width=270, justify='left', font=my_font2)

        # Put the receipt number on each canvas
        cv.create_text(95, 90, text=f'PEG{today}-{hr}{mn}{sc}', anchor='w', width=270, justify='left', font=my_font2)

        # Put the page number on each canvas
        cv.create_text(95, 110, text=f'{idx+1} of {len(canvas_lst)}', anchor='w', width=270, justify='left', font=my_font2)
        
        # Put the page header on each canvas
        for i in range(len(header_labels)):
            cv.create_text(10, i*20+30, text=header_labels[i], anchor='w', width=300, justify='left', font=my_font)
        
        # Put company information on canvas
        cv.create_text(790, 85, text='Cel: +504 9799-2662', anchor='e', justify='right', font=my_font)
        cv.create_text(790, 100, text='Email: ventas@peginservice.com', anchor='e', justify='right', font=my_font)
        cv.create_text(790, 115, text='Col. Miraflores, Calle Guanaja # 1884, Tegucigalpa, Honduras', anchor='e', justify='right', font=my_font)
        # Put the table header on each canvas
        for i in range(len(heading_labels)):
            cv.create_text(heading_labels[i][1], 155, text=heading_labels[i][0], anchor='w', width=100, justify='center', font=my_font)
        cv.create_line(10, 170, 790, 170, capstyle='round')
        
        # Fill out the table for each canvas
        for i in range(len(sub_lst)):
            cv.create_text(10, 180+i*30, text=sub_lst[i][0], anchor='w', justify='left', width=70, fill='black', font=my_font2)
            cv.create_text(60, 180+i*30, text=sub_lst[i][1], anchor='w', justify='left', width=430, fill='black', font=my_font2)
            cv.create_text(500, 180+i*30, text=sub_lst[i][2], anchor='w', justify='left', width=70, fill='black', font=my_font2)
            cv.create_text(560, 180+i*30, text=f'Lps. {sub_lst[i][3]}', anchor='w', justify='left', width=70, fill='black', font=my_font2)
            cv.create_text(660, 180+i*30, text=sub_lst[i][4], anchor='w', justify='left', width=70, fill='black', font=my_font2)
        for i in range(len(cost_sub_lst)):
            cv.create_text(720, 180+i*30, text=f'Lps. {cost_sub_lst[i]}', anchor='w', justify='left', width=70, fill='black', font=my_font2)
        if idx == len(canvas_lst)-1:
            cv.create_text(10, 180+len(sub_lst)*30, text="Resumen de Costos: ", anchor='w', justify='left', width=150, fill='black', font=my_font)
            line = cv.create_line(10, 190+len(sub_lst)*30, 790, 190+len(sub_lst)*30, capstyle='butt')
            for i in range(len(costs_labels)):
                cv.create_text(630, 210+len(sub_lst)*30+i*20, text=costs_labels[i], anchor='e', justify='right', width=300, fill='black', font=my_font)
                cv.create_text(670, 210+len(sub_lst)*30+i*20, text=f'Lps. {resumen_costos[i]}', anchor='w', justify='left', width=150, fill='black', font=my_font)
                if i==0:
                    cv.create_line(20, 310+len(sub_lst)*30+i*20, 150, 310+len(sub_lst)*30+i*20)
                    cv.create_text(20, 320+len(sub_lst)*30+i*20, text="Ing. Cesar Parada", anchor='w', justify='left', width=100, fill='black', font=my_font)
                    cv.create_image(80, 250+len(sub_lst)*30+i*20, image=signature_gif)
                    cv.create_image(250, 250+len(sub_lst)*30+i*20, image=stamp_gif)

# VARIABLES
page_cap = 1000
today = dt.datetime.today().date()
code = StringVar()
mat = StringVar()
price = StringVar()
unit = StringVar()
discount = StringVar()
discount.set('0.0')
my_font = Font(
    family = 'Times',
    size = 9,
    weight = 'bold',
    slant = 'roman',
    underline = 0,
    overstrike = 0
)
my_font2 = Font(
    family = 'Times',
    size = 9,
    weight = 'normal',
    slant = 'roman',
    underline = 0,
    overstrike = 0
)

# FIRST WINDOW WIDGETS
path_to_usr = os.path.expanduser('~')
path_to_program = os.getcwd()
logo_gif = PhotoImage(file='images/peginservice.gif')
signature_gif = PhotoImage(file='images/firma.gif')
stamp_gif = PhotoImage(file='images/new_stamp.gif', palette=1)
logo_lb = ttk.Label(content, image=logo_gif, relief='ridge') #relief: flat, groove, raised, ridge, solid, or sunken
titlelbl = ttk.Label(content, text='GENERADOR DE COTIZACIONES', justify='center', font=("Times", 16)) 
codelbl = ttk.Label(content, text='Codigo:')
matlbl = ttk.Label(content, text='Material:')
pricelbl = ttk.Label(content, text='Precio unidad:')
unitlbl = ttk.Label(content, text='Unidad:')
discount_lbl = ttk.Label(content, text='Descuento %:')

codeEntry = ttk.Entry(content, textvariable=code)
matEntry = ttk.Entry(content, textvariable=mat)
unitEntry = ttk.Entry(content, textvariable=unit)
priceEntry = ttk.Entry(content, textvariable=price)
discountEntry = ttk.Entry(content, textvariable=discount)

Add = ttk.Button(content, text='Agregar', command=add_toDb, width=25)
Remove = ttk.Button(content, text='Eliminar', command=remove_fromDb, width=25)
AddtoBudget = ttk.Button(content, text='Agregar a Cotizacion', command=add_toBudget, width=25)
RemovefromBudget = ttk.Button(content, text='Eliminar de Cotizacion', command=remove_fromBudget, width=25)
Generate = ttk.Button(content, text='Generar', command=generate_Budget, width=25)

cols = ['Código', 'Material', 'Unidad', 'Costo unidad', 'Cantidad']
tv1 = ttk.Treeview(content, columns=cols[:4], show='headings', height=15)
tv1.column(cols[0], width=15)
tv1.column(cols[1], width=150)
tv1.column(cols[2], width=15)
tv1.column(cols[3], width=15)
for col in cols[:4]:
    tv1.heading(col, text=col, command=lambda: treeview_sort_column(tv1, 'Material', False))
sb1 = ttk.Scrollbar(content, orient=VERTICAL, command=tv1.yview)
tv1.config(yscrollcommand=sb1.set)
tv1.bind("<Double-1>", update_tv1_entry)

tv2 = ttk.Treeview(content, columns=cols, show='headings', height=15)
tv2.column(cols[0], width=15)
tv2.column(cols[1], width=150)
tv2.column(cols[2], width=15)
tv2.column(cols[3], width=15)
tv2.column(cols[4], width=15)
for col2 in tv2['column']:
    tv2.heading(col2, text=col2, command=lambda: treeview_sort_column(tv2, 'Material', False))
sb2 = ttk.Scrollbar(content, orient=VERTICAL, command=tv2.yview)
tv2.config(yscrollcommand=sb2.set)

# GRID WIDGETS
titlelbl.grid(column=0, row=0, columnspan=8, padx=5, pady=5, sticky=N)
codelbl.grid(column=0, row=1, padx=5, pady=5, sticky=W)
matlbl.grid(column=0, row=2, padx=5, pady=5, sticky=W)
unitlbl.grid(column=0, row=3, padx=5, pady=5, sticky=W)
pricelbl.grid(column=0, row=4, padx=5, pady=5, sticky=W)
discount_lbl.grid(column=4, row=5, padx=5, pady=5, sticky='we')

codeEntry.grid(column=1, row=1, padx=5, pady=5,  sticky='we')
matEntry.grid(column=1, row=2, columnspan=2, padx=5, pady=5, sticky='we')
unitEntry.grid(column=1, row=3, padx=5, pady=5, sticky='we')
priceEntry.grid(column=1, row=4, padx=5, pady=5, sticky='we')
discountEntry.grid(column=5, row=5, padx=5, pady=5, sticky='we')

Add.grid(column=0, row=5, padx=5, pady=5, sticky='we')
Remove.grid(column=1, row=5, padx=5, pady=5, sticky='we')
AddtoBudget.grid(column=2, row=5, padx=5, pady=5, sticky='we')
RemovefromBudget.grid(column=6, row=5, padx=5, pady=5, sticky='we')
Generate.grid(column=7, row=5, padx=5, pady=5, sticky='we')
tv1.grid(column=0, row=6, columnspan=3, padx=5, pady=5, sticky="nsew")
tv2.grid(column=4, row=6, columnspan=4, padx=5, pady=5, sticky='nsew')
sb1.grid(column=3, row=6, padx=5, pady=5, sticky="ns")
sb2.grid(column=8, row=6, padx=5, pady=5, sticky='ns')
logo_lb.grid(column=7, row=1, rowspan=3, padx=5, pady=5, sticky='e')

# CHECK IF THERE IS A DB PRESENT
db_path = f'{path_to_usr}/Cotizaciones PegIn/database.xlsx'
isExist = os.path.exists(db_path)

if isExist:
    wb = openpyxl.load_workbook(db_path)
    ws = wb.active
    fill_treeview()
else:
    os.makedirs(f'{path_to_usr}/Cotizaciones PegIn')
    wb = Workbook()
    ws = wb.active

tv2.bind("<Double-1>", update_num_units)

root.protocol('WM_DELETE_WINDOW', on_closing)
root.mainloop()
