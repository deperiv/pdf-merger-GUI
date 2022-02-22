import glob

from tkinter        import *
from tkinter.font   import *
from tkinter        import ttk
from tkinter        import filedialog
from PyPDF2         import PdfFileMerger
from pdf_merger     import *

# GLOBAL VARIABLES
global repeated_orders

global diagnostic_folder
global order_folder
global merged_folder

global diagnostic_list_paths
global order_list_paths
global excel_path

global traceability
global cleansed_diagnostics
global cleansed_orders
global data_c

global umb_ord
global umb_ord_pdf
global umb_pdf_exc

# COSTANTS 

FRAME_WIDTH = 1334

root = Tk()
root.title('PDF Merger')
root.geometry('{}x{}'.format(FRAME_WIDTH+20, 730))

#Starting values

apply_changes_pressed = False


# Data

repeated_orders = [['', '', ''],
                   ['', '', ''],
                   ['', '', '']]

removed_orders = [[''], [''], ['']]

data_metrics_order_pdf = [["Coincidencias detectadas",''],
                         ["Coincidencia menor", ''],
                         ["Nombre en pdf", ''],
                         ["Nombre en orden", ''],
                         ["Coincidencia global" , '']]

data_metrics__pdf_excel = [["Coincidencias detectadas",''],
                            ["Coincidencia menor", ''],
                            ["Nombre en pdf", ''],
                            ["Nombre en orden", ''],
                            ["Coincidencia global" , '']]


data_ls = [['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],
           ['', ''],]

headers = [['PDF', 'ORDEN']]

data_missing = [['Fred', str(4453)],
                ['Albert', str(5624)],
                ['Mary', str(4562)],
                ['Greg', str(4522)],
                ['Candance', str(4782)],
                ['Archie', str(5658)]]


class Table(Frame):
    def __init__(self, frame, data):
        n_cols = len(data[0])
        n_rows = len(data) 
        for i in range(n_rows):
            for j in range(n_cols):
                self.e = Entry(frame, width=30)
                self.e.grid(row=i, column=j, sticky=NSEW)
                self.e.insert(END, data[i][j])

class Scroll_Table(Frame):
    def __init__(self, frame, data, headers=None, height = 90, width = 300, col_size = None):
        self.headers = headers
        if self.headers == None:
            self.data_head = data
        else:
            self.data_head = self.headers + data
        self.col_size = col_size
        self.frame = frame
        self.n_cols = len(self.data_head[0])        
        self.n_rows = len(self.data_head)
        self.height = height
        self.width = width
        self.cell_width = 50//self.n_cols

        self.my_canvas = Canvas(self.frame, height=self.height, width=self.width)
        self.my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

        self.my_scrollbar = ttk.Scrollbar(self.frame, orient=VERTICAL, command=self.my_canvas.yview)
        self.my_scrollbar.pack(side=RIGHT, fill=Y)

        self.my_canvas.configure(yscrollcommand=self.my_scrollbar.set)
        self.my_canvas.bind('<Configure>', lambda e: self.my_canvas.configure(scrollregion=self.my_canvas.bbox("all")))

        self.second_frame = Frame(self.my_canvas)

        self.my_canvas.create_window((0,0), window=self.second_frame, anchor="nw")

        for i in range(self.n_rows):
            for j in range(self.n_cols):
                if col_size == None:
                    self.e = Entry(self.second_frame, width=self.cell_width)
                else:
                    self.e = Entry(self.second_frame, width=col_size[j])
                self.e.grid(row=i, column=j, sticky=NSEW)
                self.e.insert(END, self.data_head[i][j])
    
    def update(self, data):
        if self.headers == None:
            self.data_head = data
        else:
            self.data_head = self.headers + data
        self.n_cols = len(self.data_head[0])        
        self.n_rows = len(self.data_head)
        self.cell_width = 50//self.n_cols

        for i in range(self.n_rows):
            for j in range(self.n_cols):
                if self.col_size == None:
                    self.e = Entry(self.second_frame, width=self.cell_width)
                else:
                    self.e = Entry(self.second_frame, width=self.col_size[j])
                self.e.grid(row=i, column=j, sticky=NSEW)
                self.e.insert(END, self.data_head[i][j])
               
class Examine_button(Button):
    def __init__(self, frame, command):
        self.button = Button(frame, text="Examinar", borderwidth=5, padx=20, font = font_3, command=command)
        self.button.grid(row=1,column=0)

class Error_label(Label):
    def __init__(self, frame, text):
        self.label = Label(frame, text=text, pady=0, padx=25, bg='red', foreground='white')
        self.label.grid(row=0, column=0, sticky=W)

#Setting up some fonts
font_1 = Font(family="Franklin Gothic Medium", size=16, weight="bold")
font_2 = Font(family="Arial", size=10, weight="bold")
font_3 = Font(family="Franklin Gothic Medium", size=11)

main_frame = Frame(root, bg="white", padx=10, pady=10)
main_frame.pack()

# Create frames
frame_1 = LabelFrame(main_frame, width=FRAME_WIDTH, height=60, padx=5, pady=5, borderwidth=3)
frame_2 = LabelFrame(main_frame, width=FRAME_WIDTH, height=100, padx=5, pady=5, borderwidth=3)
frame_3 = LabelFrame(main_frame, width=FRAME_WIDTH, height=110, padx=10, pady=5, borderwidth=3)
frame_4 = LabelFrame(main_frame, width=FRAME_WIDTH, height=350, padx=10, pady=5, borderwidth=3)
frame_5 = LabelFrame(main_frame, width=FRAME_WIDTH, height=90, padx=10, pady=5, borderwidth=3)

#Place frames
frame_1.grid(row=0, column= 0)
frame_2.grid(row=1, column= 0)
frame_3.grid(row=2, column= 0)
frame_4.grid(row=3, column= 0)
frame_5.grid(row=4, column= 0)

#Widgets for frame 1
frame_1.grid_propagate(0)
title = Label(frame_1, text="CONSOLIDADOR DE FACTURAS")
title.configure(font=font_1)
title.place(relx=0.4, rely=0.15)

#Widgets for frame 2

label_file_explorer = Label(frame_1, text="Hello there")
label_file_explorer.grid(row=0, column=0)



umb_ord = 0.92
umb_ord_pdf = 0.96
umb_pdf_exc = 0.93

def browseFolder_diagnostics():
    global diagnostic_folder
    diagnostic_folder = filedialog.askdirectory(initialdir = "/Desktop", title = "Select a Folder")
    label_file_explorer.configure(text="Folder Selected: "+diagnostic_folder)
    #diagnostic_list_paths = sorted(glob.glob(diagnostic_folder + '/*.pdf' ))
    
def browseFolder_orders():
    global order_folder
    order_folder = filedialog.askdirectory(initialdir = "/Desktop", title = "Select a Folder")
    label_file_explorer.configure(text="Folder Selected: "+order_folder)
    #order_list_paths = sorted(glob.glob(order_folder + '/*.pdf' ))

def browse_Excel():
    global excel_path
    excel_path = filedialog.askopenfilename(initialdir = "/Desktop", filetypes=(('Excel files', '*.xlsx'),
                                                                                ('All files', '*.*')))
    label_file_explorer.configure(text="File Selected: "+ excel_path)
    print(excel_path)

def browseFolder_merged():
    global merged_folder
    merged_folder = filedialog.askdirectory(initialdir = "/Desktop", title = "Select a Folder")
    label_file_explorer.configure(text="Folder Selected: "+ merged_folder)
    print(merged_folder)
    #diagnostic_list_paths = sorted(glob.glob(diagnostic_folder + '/*.pdf' ))

def generate_matching():
    global diagnostic_list_paths
    global order_list_paths
    global excel_path
    global traceability
    global cleansed_diagnostics
    global cleansed_orders
    global repeated_orders
    global umb_ord
    global umb_ord_pdf
    global umb_pdf_exc
    global data_c

    try:
        excel_path
        order_folder
        diagnostic_folder
    except NameError:
        error_label = Error_label(frame_1, text="ERROR: Seleccionar las carpetas y el archivo Excel antes de generar")
    else:
        error_label = Frame(frame_1, width=500, height=20, padx=25)
        error_label.grid(row=0, column=0, sticky=W)
        error_label.grid_propagate(0)

        diagnostic_list_paths = sorted(glob.glob(diagnostic_folder + '/*.pdf' ))
        order_list_paths = sorted(glob.glob(order_folder + '/*.pdf' ))

        traceability = get_traceability(excel_path)
        cleansed_diagnostics = cleanse(diagnostic_list_paths, diagnostic_folder, order_folder)
        cleansed_orders = cleanse(order_list_paths, diagnostic_folder, order_folder)

        print(cleansed_diagnostics)
        print(cleansed_orders)

        n_diagnostics = len(cleansed_diagnostics)
        n_orders = len(cleansed_orders)
        n_names_excel = len(traceability)

        label_2_4_detected.config(text=str(n_diagnostics) + " Archivos en la carpeta de pdfs\n" + str(n_orders) + " Archivos en la carpeta de órdenes\n" + str(n_names_excel) + " Facturas en el archivo Excel")

        diagnostics_matrix = get_ngram_matrix(cleansed_diagnostics, 2)
        orders_matrix = get_ngram_matrix(cleansed_orders, 2)

        repeated_orders, orders_remove  = search_multiple_orders(cleansed_orders, orders_matrix, umb_ord)
        table_4_1_2.update(repeated_orders)
        
        cleansed_orders, order_list_paths, removed_ls = remove_multiple_orders(orders_remove, cleansed_orders, order_list_paths)
        table_4_1_4.update(removed_ls)

        orders_matrix = get_ngram_matrix(cleansed_orders, 2)

        label_4_1_5.config(text="Número de órdenes a remover: " + str(len(removed_ls)))

        data_c, data_metrics_order_pdf, pdf_order_matching  = get_pdf_order_pairs(cleansed_diagnostics, cleansed_orders, diagnostics_matrix, orders_matrix, umb_ord_pdf)
        table_4_2_2.update(data_metrics_order_pdf)
        table_4_2_4.update(pdf_order_matching)

        bill_n, data_metrics_pdf_excel, pdf_excel_matching, indexes = get_pdf_excel_pairs(traceability, data_c, umb_pdf_exc)
        data_c['Billing'] = list(bill_n[np.array(indexes)[:,1]])
    
        table_4_3_2.update(data_metrics_pdf_excel)
        table_4_3_4.update(pdf_excel_matching)

        missing_billings = get_missing(indexes, traceability)
        table_4_4_3.update(missing_billings)

        label_4_4_4.config(text="Número de facturas a consolidar\nmanualmente: " + str(len(missing_billings)))

def merge_files():
    pdf_index = data_c['PDF_index']
    order_index = data_c['Order_index']
    billing_ls = data_c['Billing']
    label_5_3_1.config(text="ESTADO: EN PROCESO")
    for i in range(len(pdf_index)):
        merger = PdfFileMerger(strict=False)
        merger.append(diagnostic_list_paths[pdf_index[i]])
        merger.append(order_list_paths[order_index[i]])
        merger.write(merged_folder + '/FE' + str(billing_ls[i])+'.pdf')
        merger.close()
        progress_bar_5_3_2['value'] = (i+1)*100//len(pdf_index)
        label_5_3_3.config(text="Facturas consolidadas: " + str(i+1) + " de " + str(len(pdf_index)))
        root.update_idletasks()
    label_5_3_1.config(text="ESTADO: FINALIZADO")

frame_2.grid_propagate(0)
frame_2_1 = Frame(frame_2, padx=25, pady=0)
frame_2_1.place(relx=0.1)
label_2_1 = Label(frame_2_1, text="Seleccionar carpeta con pdfs", pady=5, padx=5, font=font_2)
label_2_1.grid(row=0,column=0)
button_2_1 = Examine_button(frame_2_1, command=browseFolder_diagnostics)

frame_2_2 = Frame(frame_2, padx=25, pady=0)
frame_2_2.place(relx=0.3)
label_2_2 = Label(frame_2_2, text="Seleccionar carpeta con órdenes", pady=5, padx=5, font=font_2)
label_2_2.grid(row=0,column=0)
button_2_2 = Examine_button(frame_2_2, command=browseFolder_orders)

frame_2_3 = Frame(frame_2, padx=25, pady=0)
frame_2_3.place(relx=0.5)
label_2_3 = Label(frame_2_3, text="Seleccionar archivo con lista Excel", pady=5, padx=5, font=font_2)
label_2_3.grid(row=0,column=0)
button_2_3 = Examine_button(frame_2_3, command=browse_Excel)

var_a = 4
frame_2_4 = LabelFrame(frame_2, padx=15, pady=5, borderwidth=3)
frame_2_4.place(relx=0.81)
label_2_4 = Label(frame_2_4, text="Se han detectado:", font=font_2)
label_2_4_detected = Label(frame_2_4, text=str(var_a) + " Archivos en la carpeta de pdfs\n" + str(var_a) + " Archivos en la carpeta de órdenes\n" + str(var_a) + " Facturas en el archivo Excel", justify=LEFT)
label_2_4.grid(row=0,column=0)
label_2_4_detected.grid(sticky = W, row=1,column=0)

#Widgets for frame 3
frame_3.grid_propagate(0)
label_3 = Label(frame_3, text="CONFIGURACIÓN DE UMBRALES", pady=0, padx=25, font=font_2)
label_3.grid(row=0,column=0, sticky=W)

frame_3_1 = Frame(frame_3, padx=10)
frame_3_1.place(relx=0.03, rely=0.35)
label_3_1 = Label(frame_3_1, text="Similitud órdenes repetidas", padx=5, justify=LEFT)
label_3_1.grid(row=0,column=0)

frame_3_1_1 = Frame(frame_3_1, pady=5)
frame_3_1_1.grid(row=1,column=0)
input_3_1 = Entry(frame_3_1_1, width=5, borderwidth=2)
input_3_1.grid(row=0,column=0)
input_3_1.insert(0, '0.93')
arrow_3_1 = Label(frame_3_1_1, text="   ===>", font = font_2)
arrow_3_1.grid(row=0,column=1)
value_3_1 = Label(frame_3_1_1, text=str(umb_ord), padx=15)
value_3_1.grid(row=0,column=2)

frame_3_2 = Frame(frame_3, padx=10)
frame_3_2.place(relx=0.18, rely=0.35)
label_3_2 = Label(frame_3_2, text="Similitud órdenes y pdfs", padx=5, justify=LEFT)
label_3_2.grid(row=0,column=0,sticky=W)

frame_3_2_1 = Frame(frame_3_2, pady=5)
frame_3_2_1.grid(row=1,column=0)
input_3_2 = Entry(frame_3_2_1, width=5, borderwidth=2)
input_3_2.grid(row=0,column=0)
input_3_2.insert(0, '0.91')
arrow_3_2 = Label(frame_3_2_1, text="   ===>", font = font_2)
arrow_3_2.grid(row=0,column=1)
value_3_2 = Label(frame_3_2_1, text=str(umb_ord_pdf), padx=15)
value_3_2.grid(row=0,column=2)

frame_3_3 = Frame(frame_3, padx=10)
frame_3_3.place(relx=0.32, rely=0.35)
label_3_3 = Label(frame_3_3, text="Similitud pdfs y lista Excel", padx=5, justify=LEFT)
label_3_3.grid(row=0,column=0)

frame_3_3_1 = Frame(frame_3_3, pady=5)
frame_3_3_1.grid(row=1,column=0)
input_3_3 = Entry(frame_3_3_1, width=5, borderwidth=2)
input_3_3.grid(row=0,column=0)
input_3_3.insert(0, '0.90')
arrow_3_3 = Label(frame_3_3_1, text="   ===>", font = font_2)
arrow_3_3.grid(row=0,column=1)
value_3_3 = Label(frame_3_3_1, text=str(umb_pdf_exc), padx=15)
value_3_3.grid(row=0,column=2)

def apply_changes():    
    global umb_ord
    global umb_ord_pdf
    global umb_pdf_exc
    apply_changes_pressed = True
    umb_ord = float(input_3_1.get())
    umb_ord_pdf = float(input_3_2.get())
    umb_pdf_exc = float(input_3_3.get())
    if umb_ord > 1 or umb_ord_pdf > 1 or umb_pdf_exc > 1:
        error_label = Error_label(frame_1, text="ERROR: Introducir valores entre 0 y 1")
    else: 
        error_label = Frame(frame_1, width=300, height=20, padx=25)
        error_label.grid(row=0, column=0, sticky=W)
        error_label.grid_propagate(0)
        value_3_1.configure(text=str(input_3_1.get()))
        value_3_2.configure(text=str(input_3_2.get()))
        value_3_3.configure(text=str(input_3_3.get()))

def set_default_values():
    error_label_3 = Frame(frame_1, width=300, height=20, pady=0, padx=25)
    error_label_3.grid(row=0,column=0, sticky=W)    
    umb_ord = 0.93
    umb_ord_pdf = 0.91
    umb_pdf_exc = 0.90
    value_3_1.configure(text=umb_ord)
    value_3_2.configure(text=umb_ord_pdf)
    value_3_3.configure(text=umb_pdf_exc)

frame_3_4 = Frame(frame_3, padx=10, pady=0)
frame_3_4.place(relx=0.48, rely=0.2)
button_3_4 = Button(frame_3_4, text="Aplicar cambios", padx=5, pady=14, borderwidth=5, font=font_2, command=apply_changes)
button_3_4.grid(row=0,column=0)

frame_3_5 = Frame(frame_3, padx=10, pady=0)
frame_3_5.place(relx=0.6, rely=0.2)
button_3_5 = Button(frame_3_5, text="Establecer valores por defecto", padx=5, pady=14, borderwidth=5, font=font_2, command=set_default_values)
button_3_5.grid(row=0,column=0)

frame_3_6 = LabelFrame(frame_3, padx=7, pady=7, borderwidth=3)
frame_3_6.place(relx=0.81, rely=0.1)
button_3_6 = Button(frame_3_6, text="GENERAR COINCIDENCIAS", padx=15, pady=15, borderwidth=5, font=font_2, command=generate_matching)
button_3_6.grid(row=0,column=0)

#Widgets for frame 4
frame_4.grid_propagate(0)

width_frames_4 = 355
height_frames_4 = 338
width_table = 300

#ÓRDENES REPETIDAS
frame_4_1 = LabelFrame(frame_4, padx=10, pady=5, borderwidth=5, height=height_frames_4, width=width_frames_4)
frame_4_1.grid_propagate(0)
frame_4_1.grid(row=0,column=0, sticky=N)
title_4_1 = Label(frame_4_1, text="ÓRDENES REPETIDAS")
title_4_1.grid(row=0, sticky=W)
title_4_1.configure(font=font_2)
label_4_1_1 = Label(frame_4_1, text="Órdenes repetidas detectadas")
label_4_1_1.grid(row=1, sticky=W)
frame_4_1_2 = Frame(frame_4_1, padx=5, pady=5)
frame_4_1_2.grid(row=2)
table_4_1_2 = Scroll_Table(frame_4_1_2, repeated_orders, [['NOMBRE 1', 'NOMBRE 2', '%']], width=width_table, col_size=[20,20,5])
label_4_1_3 = Label(frame_4_1, text="Las siguientes órdenes serán removidas del proceso\nde consolidación", justify=LEFT)
label_4_1_3.grid(row=3, sticky=W)
frame_4_1_4 = Frame(frame_4_1, padx=5, pady=5)
frame_4_1_4.grid(row=4)
table_4_1_4 = Scroll_Table(frame_4_1_4, removed_orders, [['ORDEN']], width=width_table)
label_4_1_5 = Label(frame_4_1, text="Número de órdenes a remover: " + str(5))
label_4_1_5.place(rely=0.92, relx=0.25)

#COINCIDENCIAS ENTRE PDFS Y ÓRDENES
frame_4_2 = LabelFrame(frame_4, padx=10, pady=5, borderwidth=5, height=height_frames_4, width=width_frames_4)
frame_4_2.grid_propagate(0)
frame_4_2.grid(row=0, column=1, sticky=N)
title_4_2 = Label(frame_4_2, text="COINCIDENCIAS ENTRE PDFS Y ÓRDENES")
title_4_2.grid(row=0, sticky=W)
title_4_2.configure(font=font_2) 

label_4_2_1 = Label(frame_4_2, text="TABLA DE MÉTRICAS", pady=3)
label_4_2_1.grid(row=1, sticky=W)
frame_4_2_2 = Frame(frame_4_2, padx=5, pady=5)
frame_4_2_2.grid(row=2)
table_4_2_2 = Scroll_Table(frame_4_2_2, data_metrics_order_pdf, width=width_table)
label_4_2_3 = Label(frame_4_2, text="TABLA DE COINCIDENCIAS", pady=3)
label_4_2_3.grid(row=3, sticky=W)
frame_4_2_4 = Frame(frame_4_2, padx=5, pady=5)
frame_4_2_4.grid(row=4)
table_4_2_4 = Scroll_Table(frame_4_2_4, data_ls, headers, width=width_table, height=120)

#COINCIDENCIAS ENTRE PDFS Y EXCEL
frame_4_3 = LabelFrame(frame_4, padx=10, pady=5, borderwidth=5, height=height_frames_4, width=width_frames_4)
frame_4_3.grid_propagate(0)
frame_4_3.grid(row=0, column=2, sticky=N)
title_4_3 = Label(frame_4_3, text="COINCIDENCIAS ENTRE PDFS Y EXCEL")
title_4_3.grid(row=0, sticky=W)
title_4_3.configure(font=font_2)

label_4_3_1 = Label(frame_4_3, text="TABLA DE MÉTRICAS", pady=3)
label_4_3_1.grid(row=1, sticky=W)

frame_4_3_2 = Frame(frame_4_3, padx=5, pady=5)
frame_4_3_2.grid(row=2)
table_4_3_2 = Scroll_Table(frame_4_3_2, data_metrics__pdf_excel, width=width_table)

label_4_3_3 = Label(frame_4_3, text="TABLA DE COINCIDENCIAS", pady=3)
label_4_3_3.grid(row=3, sticky=W)

frame_4_3_4 = Frame(frame_4_3, padx=5, pady=5)
frame_4_3_4.grid(row=4)
table_4_3_4 = Scroll_Table(frame_4_3_4, data_ls, headers = [['PDF', 'EXCEL']], width=width_table, height=120)

#ARCHIVOS FALTANTES
frame_4_4 = LabelFrame(frame_4, padx=10, pady=5, borderwidth=5, height=height_frames_4, width=240)
frame_4_4.grid_propagate(0)
frame_4_4.grid(row=0, column=4, sticky=N)
title_4_4 = Label(frame_4_4, text="FACTURAS FALTANTES")
title_4_4.grid(row=0, sticky=W)
title_4_4.configure(font=font_2)
label_4_4_1 = Label(frame_4_4, text="TABLA DE FACTURAS FALTANTES", pady=10)
label_4_4_1.grid(row=1, sticky=W)

label_4_4_2 = Label(frame_4_4, text="Las siguientes facturas necesitan \nconsolidarse manualmente", justify=LEFT, pady=10)
label_4_4_2.grid(row=2, sticky=W)


frame_4_4_3 = Frame(frame_4_4, pady=10)
frame_4_4_3.grid(row=3)
table_4_4_3 = Scroll_Table(frame_4_4_3, data_missing, width=width_table-110, headers=[['NOMBRES', '# FACTURA']], col_size=[20, 12])

label_4_4_4 = Label(frame_4_4, pady=3, text="Número de facturas a consolidar\nmanualmente: ")
label_4_4_4.grid(row=4)
button_4_4_4 = Button(frame_4_4, text="Exportar archivo Excel", borderwidth=5, padx=20)
button_4_4_4.place(rely=0.85, relx=0.10)

#Widgets for frame 5
frame_5.grid_propagate(0)

dev_name = Label(frame_1, text="Developed by Daniel Peraza Rivera", padx=5, pady=5, font=font_2)
dev_name.place(relx=0.8, rely=0.2)

frame_5_1 = Frame(frame_5, padx=20)
frame_5_1.grid(row=0, column=0, sticky=N)
label_5_1_1 = Label(frame_5_1, text="Seleccionar carpeta para guardar consolidados", padx=5, pady=5)
label_5_1_1.grid(row=0)
button_5_1_2 = Examine_button(frame_5_1, command=browseFolder_merged)

frame_5_2 = LabelFrame(frame_5, padx=15, pady=10, borderwidth=3)
frame_5_2.grid(row=0, column=1, sticky=N)
button_5_2 = Button(frame_5_2, text="JUNTAR Y GUARDAR CONSOLIDADOS", borderwidth=5, font=font_2, padx=5, pady=10, command=merge_files)
button_5_2.grid(row=0)

frame_5_3 = LabelFrame(frame_5, padx=20, pady=2, borderwidth=3)
frame_5_3.grid(row = 0, column=2, sticky=N)
label_5_3_1 = Label(frame_5_3, text="ESTADO: SIN INICIAR " , pady=2)
label_5_3_1.grid(row=0, sticky=W)
progress_bar_5_3_2 = ttk.Progressbar(frame_5_3, orient=HORIZONTAL, length=500, mode='determinate')
progress_bar_5_3_2.grid(row=1)

label_5_3_3 = Label(frame_5_3, text ="Facturas consolidadas: " + '-' + " de " + '-', pady=2)
label_5_3_3.grid(row=2, sticky=W)

frame_5_4 = LabelFrame(frame_5, padx=10, pady=10, borderwidth=3)
frame_5_4.grid(row = 0, column=3, sticky=N)
button_5_4 = Button(frame_5_4, text="REINICIAR TODO", borderwidth=5, font=font_2, padx=5, pady=10)
button_5_4.grid(row=0)

root.mainloop()
