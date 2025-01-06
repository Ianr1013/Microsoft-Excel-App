import tkinter as tk
from tkinter import ttk
import openpyxl

def toggle_mode():
    if mode.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def load_data():
    path = "C:\SPAREZONE\ExcelApp\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    # print(list_values)
    for col_name in list_values[0]:
        # putting the column names to the treeview
        view.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        view.insert('', tk.END, values=value_tuple)

def insert_row():
    #insert row into excel excel sheet and treeview
    name = name_entry.get()
    age = int(age_spinbox.get())
    sub = combobox.get()
    emp_stats = "Employed" if a.get() else "Unemployed"

    #for the excel sheet
    path = "C:\SPAREZONE\ExcelApp\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, sub, emp_stats]
    sheet.append(row_values)
    workbook.save(path)

    #for the treeview
    view.insert('', tk.END, values=row_values)

    #clear values once clicking the button
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    combobox.set(combo_list[0])
    checkbutton.state(["!selected"])



combo_list = ["Subscribed", "Not Subscribed", "Other"]

root = tk.Tk()
style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(root)
frame.pack()

widget_frame = ttk.LabelFrame(frame, text="Insert Row")
widget_frame.grid(row=0, column=0, padx=20, pady=10)

#entry label for the user to insert his/her name
name_entry = ttk.Entry(widget_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0,5), sticky="ew")

#creating the age spinbox
age_spinbox = ttk.Spinbox(widget_frame, from_ = 18, to = 100)
age_spinbox.insert(0, "Age")
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

#subscription menu
combobox = ttk.Combobox(widget_frame, values=combo_list)
combobox.current(0)
combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

#empployment checkbox
a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widget_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0,padx=5, pady=5, sticky="nsew")

button = ttk.Button(widget_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

sep = ttk.Separator(widget_frame)
sep.grid(row=5, column=0, padx=(20,10), pady=10, sticky="ew")

mode = ttk.Checkbutton(widget_frame, text="Mode", style="Switch", command=toggle_mode)
mode.grid(row=5, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
scroll = ttk.Scrollbar(treeFrame)
scroll.pack(side="right", fill="y")

cols = ("Name", "Age", "Subscription", "Employment")
view = ttk.Treeview(treeFrame, show="headings", yscrollcommand=scroll.set, columns=cols, height=13)
view.column("Name", width=100)
view.column("Age", width=50)
view.column("Subscription", width=100)
view.column("Employment", width=100)
view.pack()
scroll.config(command=view.yview)
load_data()




root.mainloop()