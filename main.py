import math
import os
import sys
import threading
import tkinter as tk
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import messagebox as mb
import numpy as np
import pandas as pd
import tksheet
from tkscrolledframe import ScrolledFrame

ddta = list
opened = False
saved = False
counted = False
scale_factor = 1

root = tk.Tk()


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


#Функция загрузки основного окна
def download_root(prog_root):
    prog_root.resizable(True, True)
    prog_root.geometry("1280x720")
    prog_root.title("")
    icon = PhotoImage(file=resource_path("static/icon.png"))
    prog_root.iconphoto(False, icon)
    prog_root.configure(bg="#fcfcfc")

thread0 = threading.Thread(target=download_root(root))
thread0.start()


def truncate(number, digits) -> float:
    print(number)
    nb_decimals = len(str(number).split('.')[1])
    if nb_decimals <= digits:
        return number
    stepper = 10.0 ** digits
    return math.trunc(stepper * number) / stepper


#Функция выбора файла
def select_file(event=""):
    filetypes = (('excel files', ('*.xlsx', '*.xls', '*.ods')), ('text files', '*.txt'), ('All files', '*.*'))
    global filename
    filename = fd.askopenfilename(title='Открыть файл', filetypes=filetypes)
    try:
        data = pd.read_table(filename, sep='\t', header=None)
    except:
        data = pd.read_excel(filename, header=None)
    global a
    global DATA
    DATA = data
    try:
        a = np.zeros((data.shape[0] - 1, data.shape[1] - 1))
        for i in range(data.shape[0] - 1):
            for j in range(data.shape[1] - 1):
                a[i][j] = data.iat[i + 1, j + 1]
        rd = a.max()
        for i in range(data.shape[0] - 1):
            for j in range(data.shape[1] - 1):
                a[i][j] = a[i][j] / rd
        sheet.set_sheet_data([[f"{data.iat[ri, cj]}" for cj in range(data.shape[1])] for ri in range(data.shape[0])])
        sheet.enable_bindings("all")
        sheet.bind("<Control-C>", sheet.copy)
        sheet.bind("<Control-X>", sheet.cut)
        sheet.bind("<Control-V>", sheet.paste)
        sheet.bind("<Delete>", sheet.delete)
        sheet.set_all_cell_sizes_to_text()
        sheet.set_all_column_widths(width=70)
        global ddta
        ddta = pd.DataFrame(sheet.get_sheet_data())
        ddta.convert_dtypes()
        ddta = ddta.to_dict()
        global x
        x = []
        data = a
        global z, sumarr
        z = np.ones((data.shape[1], 1))
        x = np.ones((data.shape[0], 1))
        for l in range(3):
            for i in range(data.shape[0]):
                sum = 0
                for j in range(data.shape[1]):
                    sum += data[i][j] * z[j]
                x[i] = sum
            sumi = 0
            for i in z:
                sumi += i
            sumarr = sumi
            for i in range(data.shape[0]):
                x[i] = x[i] / sumarr
            for i in range(data.shape[1]):
                sum = 0
                for j in range(data.shape[0]):
                    sum += (1 - data[j][i]) * x[j]
                z[i] = sum
            sumi = 0
            for i in x:
                sumi += i
            sumarr = sumi
            for i in range(data.shape[1]):
                z[i] = z[i] / sumarr
        for i in range(data.shape[0]):
            sum = 0
            for j in range(data.shape[1]):
                sum += data[i][j] * z[j]
            x[i] = sum
        sumi = 0
        for i in x:
            sumi += i
        sumarr = sumi
        SUM = float('-inf')
        for i in range(data.shape[1]):
            sum = 0
            for j in range(data.shape[0]):
                sum += x[j]
                if sum > SUM:
                    SUM = sum
        for i in range(data.shape[1]):
            sum = 0
            for j in range(data.shape[0]):
                sum += (1 - data[j][i]) * x[j]
            if SUM != 0:
                z[i] = sum / (SUM / 2)
            else:
                z[i] = 0
        for i in range(data.shape[0]):
            x[i] = float(truncate(float(x[i]), 3))
        for j in range(data.shape[1]):
            z[j] = float(truncate(float(z[j]), 3))
        sheet1.set_sheet_data([[f"{z[ri][cj]}" for cj in range(z.shape[1])] for ri in range(z.shape[0])])
        sheet1.enable_bindings("all")
        sheet1.bind("<Control-C>", sheet1.copy)
        sheet1.bind("<Control-X>", sheet1.cut)
        sheet1.bind("<Control-V>", sheet1.paste)
        sheet1.bind("<Delete>", sheet1.delete)
        sheet1.set_all_cell_sizes_to_text()
        sheet1.set_all_column_widths(width=70)
        sheet2.set_sheet_data([[f"{x[ri][cj]}" for cj in range(x.shape[1])] for ri in range(x.shape[0])])
        sheet2.enable_bindings("all")
        sheet2.bind("<Control-C>", sheet2.copy)
        sheet2.bind("<Control-X>", sheet2.cut)
        sheet2.bind("<Control-V>", sheet2.paste)
        sheet2.bind("<Delete>", sheet2.delete)
        sheet2.set_all_cell_sizes_to_text()
        sheet2.set_all_column_widths(width=70)
        counted = True
        b2['state'] = 'normal'
        b3['state'] = 'normal'
    except ValueError:
        messagebox.showerror("Ошибка", "Неправильный формат данных")


#Функция пересчета весов
def count_z(event=""):
    data = pd.DataFrame(sheet.get_sheet_data())
    try:
        a = np.zeros((data.shape[0] - 1, data.shape[1] - 1))
        for i in range(data.shape[0] - 1):
            for j in range(data.shape[1] - 1):
                a[i][j] = data.iat[i + 1, j + 1]
        rd = a.max()
        for i in range(data.shape[0] - 1):
            for j in range(data.shape[1] - 1):
                a[i][j] = a[i][j] / rd
        sheet.set_sheet_data([[f"{data.iat[ri, cj]}" for cj in range(data.shape[1])] for ri in range(data.shape[0])])
        sheet.enable_bindings("all")
        sheet.bind("<Control-C>", sheet.copy)
        sheet.bind("<Control-X>", sheet.cut)
        sheet.bind("<Control-V>", sheet.paste)
        sheet.bind("<Delete>", sheet.delete)
        sheet.set_all_cell_sizes_to_text()
        sheet.set_all_column_widths(width=70)
        global x
        x = []
        data = a
        global z, sumarr
        z = np.ones((data.shape[1], 1))
        x = np.ones((data.shape[0], 1))
        for l in range(10):
            for i in range(data.shape[0]):
                sum = 0
                for j in range(data.shape[1]):
                    sum += data[i][j] * z[j]
                x[i] = sum
            sumi = 0
            for i in z:
                sumi += i
            sumarr = sumi
            for i in range(data.shape[0]):
                x[i] = x[i] / sumarr
            for i in range(data.shape[1]):
                sum = 0
                for j in range(data.shape[0]):
                    sum += (1 - data[j][i]) * x[j]
                z[i] = sum
            sumi = 0
            for i in x:
                sumi += i
            sumarr = sumi
            for i in range(data.shape[1]):
                z[i] = z[i] / sumarr
        for i in range(data.shape[0]):
            sum = 0
            for j in range(data.shape[1]):
                sum += data[i][j] * z[j]
            x[i] = sum
        sumi = 0
        for i in x:
            sumi += i
        sumarr = sumi
        SUM = -100
        for i in range(data.shape[1]):
            sum = 0
            for j in range(data.shape[0]):
                sum += x[j]
                if sum > SUM:
                    SUM = sum
        for i in range(data.shape[1]):
            sum = 0
            for j in range(data.shape[0]):
                sum += (1 - data[j][i]) * x[j]
            z[i] = sum / (SUM / 2)
        for i in range(data.shape[0]):
            x[i] = float(truncate(float(x[i]), 3))
        for j in range(data.shape[1]):
            z[j] = float(truncate(float(z[j]), 3))
        sheet1.set_sheet_data([[f"{z[ri][cj]}" for cj in range(z.shape[1])] for ri in range(z.shape[0])])
        sheet1.enable_bindings("all")
        sheet1.bind("<Control-C>", sheet1.copy)
        sheet1.bind("<Control-X>", sheet1.cut)
        sheet1.bind("<Control-V>", sheet1.paste)
        sheet1.bind("<Delete>", sheet1.delete)
        sheet1.set_all_cell_sizes_to_text()
        sheet1.set_all_column_widths(width=70)
        sheet2.set_sheet_data([[f"{x[ri][cj]}" for cj in range(x.shape[1])] for ri in range(x.shape[0])])
        sheet2.enable_bindings("all")
        sheet2.bind("<Control-C>", sheet2.copy)
        sheet2.bind("<Control-X>", sheet2.cut)
        sheet2.bind("<Control-V>", sheet2.paste)
        sheet2.bind("<Delete>", sheet2.delete)
        sheet2.set_all_cell_sizes_to_text()
        sheet2.set_all_column_widths(width=70)
        counted = True
        global ddta
        ddta = pd.DataFrame(sheet.get_sheet_data())
        ddta.convert_dtypes()
        ddta = ddta.to_dict()
    except ValueError:
        messagebox.showerror("Ошибка", "Неправильный формат данных")


#Функция создания таблицы с весами
def add_to_excel(event=""):
    data = pd.DataFrame(sheet.get_sheet_data())
    data.convert_dtypes()
    data = data.to_dict()
    print(data)
    print(ddta)
    if ddta != data:
        answer = mb.askyesno(
            title="Внимание",
            message="Веса не пересчитаны. Не хотите ли пересчитать?")
        if answer:
            count_z()
    data = sheet.get_sheet_data()
    print(data)
    for i in range(len(data) - 1):
        for j in range(len(data[i]) - 1):
            data[i + 1][j + 1] = float(data[i + 1][j + 1])
    y = ["Веса задач"]
    for i in range(z.shape[0]):
        y.append(float(z[i]))
    h = np.row_stack((y, data))
    f = ["Результаты", "Веса учеников"]
    for i in range(len(x)):
        f.append(float(x[i]))
    r = np.column_stack((f, h))
    df = pd.DataFrame(r)
    df.convert_dtypes(convert_floating=True)
   #df = df.convert_objects(convert_numeric=True)
    print(df)
    filetypes = (('excel files', '*.xlsx'), ('All files', '*.*'))
    file_name = fd.asksaveasfilename(title='Новый файл', filetypes=filetypes)
    df.to_excel(file_name + ".xlsx", index=False, header=None)
    saved = True


def download_new_table():
    DATA = sheet.get_sheet_data()
    data = pd.DataFrame(DATA)
    filetypes = (('excel files', '*.xlsx'), ('All files', '*.*'))
    file_name = fd.asksaveasfilename(title='Новый файл', filetypes=filetypes)
    data.to_excel(file_name + ".xlsx", index=False, header=None)
    saved = True


#Функция выхода из программы
def exit_program(event=""):
    if saved:
        root.destroy()
    else:
        answer = mb.askyesno(
            title="Внимание",
            message="Ваши данные не сохранены. Вы действительно хотите выйти?")
        if answer:
            root.destroy()


#Функция вызова окна со справкой
def desc():
    b = Toplevel()
    b.title("Справка")
    p1 = tk.PhotoImage(file=resource_path("static/about.png"))
    b.iconphoto(False, p1)
    b.focus()
    sf = ScrolledFrame(b, width=1280, height=720)
    inner_frame = sf.display_widget(Frame)
    img_label = tk.Label(inner_frame)
    img_label.image = tk.PhotoImage(file=resource_path("static/Vesa.png"))
    img_label['image'] = img_label.image
    img_label.pack()
    sf.pack()


#Функция разметки расположения таблиц
def download_sheet():
    global scale_factor
    sheet.height_and_width(600, 900)
    sheet.default_column_width(width=30)
    sheet.place(relx=0.03, rely=0.22)
    sheet1.height_and_width(600, 130)
    sheet1.default_column_width(width=30)
    sheet1.place(relx=0.77, rely=0.22)
    sheet2.height_and_width(600, 130)
    sheet2.default_column_width(width=30)
    sheet2.place(relx=0.88, rely=0.22)


#Функция загрузки текста для интерфейса
def download_text(prog_root):
    root.image = tk.PhotoImage(file=resource_path("static/weights_tasks.png"))
    label = tk.Label(root, image=root.image)
    label.place(relx=0.77, rely=0.18)
    root.image1 = tk.PhotoImage(file=resource_path("static/weights_pupil.png"))
    label1 = tk.Label(root, image=root.image1)
    label1.place(relx=0.88, rely=0.18)


def resize(event):
    canvas.create_line(0, 60, root.winfo_width(), 60, fill='#D6D6D6')


#main
root.bind('<Configure>', resize)
canvas = Canvas(root)
canvas.place(x=0, y=0, relheight=0.1, relwidth=1)
canvas.create_line(0, 60, root.winfo_width(), 60, fill='#D6D6D6')
sheet = tksheet.Sheet(root)
sheet1 = tksheet.Sheet(root)
sheet2 = tksheet.Sheet(root)
thread = threading.Thread(target=download_sheet())
thread.start()
tread1 = threading.Thread(target=download_text(root))
tread1.start()
i1 = PhotoImage(file=resource_path("static/choose_file.png"))
i3d = PhotoImage(file=resource_path("static/recount_weights.png"))
b1 = Button(root, command=select_file, text="Выбрать файл", bg="#fcfcfc", image=i1, borderwidth=0)
b1.place(relx=0.2, rely=0.1)
i2d = PhotoImage(file=resource_path("static/unload_table.png"))
b2 = Button(root, command=add_to_excel, text="Выгрузить табл. с весами", bg="#fcfcfc", image=i2d, borderwidth=0, state=DISABLED)
b2.place(relx=0.8, rely=0.1)
b3 = Button(root, command=count_z, text="Пересчитать веса", bg="#fcfcfc", image=i3d, borderwidth=0, state=DISABLED)
b3.place(relx=0.5, rely=0.1)
"""
main_menu = Menu(root, bg="#ffffff", fg="#000000", bd=0, activeborderwidth=0, relief=FLAT, tearoff=0)
root.config(menu=main_menu)
file_menu = Menu(main_menu, tearoff=0, bg="#ffffff", fg="#000000", takefocus=True, bd=0, activeborderwidth=0)
file_menu.add_command(label="Выбрать файл", command=select_file, underline=0, accelerator="Ctrl+O")
file_menu.add_command(label="Выгрузить табл. с весами", command=add_to_excel, underline=1, accelerator='Ctrl+S')
file_menu.add_command(label="Пересчитать веса", command=count_z, underline=2, accelerator='Ctrl+P')
file_menu.add_separator()
file_menu.add_command(label="Выход", underline=3, command=exit_program)
main_menu.add_cascade(label="Файл", underline=0, menu=file_menu)
main_menu.add_command(label="Справка", underline=0, command=desc)
file_menu.bind_all('<Control-q>', exit_program)
file_menu.bind_all('<Control-o>', select_file)
file_menu.bind_all('<Control-s>', add_to_excel)
file_menu.bind_all('<Control-p>', count_z)
"""
i4 = PhotoImage(file=resource_path("static/logo.png"))
Label(image=i4).place(x=0, y=0)
i5 = PhotoImage(file=resource_path("static/help.png"))
Button(root, command=desc, image=i5, borderwidth=0).place(x=400, y=20)
i6 = PhotoImage(file=resource_path("static/exit.png"))
Button(root, command=exit_program, image=i6, borderwidth=0).place(x=550, y=20)
root.mainloop()
