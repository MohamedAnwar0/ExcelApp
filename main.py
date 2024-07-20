import tkinter as tk
from tkinter import ttk
import openpyxl

def switching_mode():

    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


def load_data():

    path = "people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    
    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        tree.heading(col_name, text=col_name)
    for values_tuple in list_values:
        tree.insert("", tk.END, values=values_tuple)


def insert_row():
    name = name_entry.get()
    age = age_spinbox.get()
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name, age, subscription_status, employment_status)

    # Insert row into Excel sheet
    path = "./people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    tree.insert('', tk.END, values=row_values)

    # Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])


window = tk.Tk()

# theme of the window <forest-ttk-theme>
style = ttk.Style(window)
window.tk.call("source", "forest-light.tcl")
window.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")


frame = ttk.Frame(window)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=5, pady=5)

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0, 'Age')
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

combo_list = ['Subscribed', 'Not subscribed', 'Other']
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

a = tk.BooleanVar()
employed_checkbutton = ttk.Checkbutton(widgets_frame, text="Employed", variable=a)
employed_checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

# Seperator between two frames
separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

# mode switching
mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=switching_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

tree_frame = ttk.Frame(frame)
tree_frame.grid(row=0, column=1, padx=5, pady=10)

scroll = ttk.Scrollbar(tree_frame)
scroll.pack(side="right", fill="y")


cols = ("Name", "Age", "Subscription", "Employment")
tree = ttk.Treeview(tree_frame, show="headings", yscrollcommand=scroll.set,
                    columns=cols, height=13)

tree.column("Name", width=100)
tree.column("Age", width=50)
tree.column("Subscription", width=100)
tree.column("Employment", width=100)

tree.pack()
scroll.config(command=tree.yview)
load_data()


window.mainloop()
