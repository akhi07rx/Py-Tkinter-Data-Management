# test_spc

import tkinter as tk
from tkinter import ttk
import openpyxl


def theme_switch_mode():
    if theme_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")


def load_data():
    path = "C://Users//AKHI//Desktop//PY TEST//Tkinter Data//assets//xlxs//people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)

    for col_name in list_values[0]:
        tree_view.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        tree_view.insert('', tk.END, values=value_tuple)


def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "EMPLOYED" if cv.get() else "UNEMPLOYED"

    print("ADDED TO DB\n", name, age, subscription_status, employment_status)

    path = "C://Users//AKHI//Desktop//PY TEST//Tkinter Data//assets//xlxs//people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    tree_view.insert('', tk.END, values=row_values)

    name_entry.delete(0, 'end')
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, 'end')
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])
    cv.set(False)


app = tk.Tk()
app.title("TEST")
# app.config(bg="#313130")

style = ttk.Style(app)
app.tk.call("source", "Themes//forest-dark.tcl")
app.tk.call("source", "Themes//forest-light.tcl")
style.theme_use("forest-dark")


combo_list = ["SUBSCRIBED", "NOT SUBSCRIBED", "OTHERS"]

frame = ttk.Frame(app)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="INSERT ROW")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "NAME")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete(0, 'end'))
name_entry.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))

age_spinbox = ttk.Spinbox(widgets_frame, from_=18, to=100)
age_spinbox.insert(0, "AGE")
age_spinbox.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

cv = tk.BooleanVar()
check_button = ttk.Checkbutton(widgets_frame, text="EMPLOYED", variable=cv)
check_button.grid(row=3, column=0, sticky="nsew", padx=5, pady=5)

button = ttk.Button(widgets_frame, text="INSERT", command=insert_row)
button.grid(row=4, column=0, sticky="nsew", padx=5, pady=5)


separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

theme_switch = ttk.Checkbutton(
    widgets_frame, text="mode", style="Switch", command=theme_switch_mode)
theme_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")


tree_frame = ttk.Frame(frame)
tree_frame.grid(row=0, column=1, pady=10)
tree_scroll = ttk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")

cols = ["NAME", "AGE", "SUBSCRIBED", "EMPLOYMENT"]
tree_view = ttk.Treeview(tree_frame, show="headings",
                         yscrollcommand=tree_scroll.set, columns=cols, height=13)
tree_view.pack()
tree_view.column("NAME", width=100)
tree_view.column("AGE", width=50)
tree_view.column("SUBSCRIBED", width=100)
tree_view.column("EMPLOYMENT", width=100)
tree_view.pack()
tree_scroll.config(command=tree_view.yview)

load_data()


app.mainloop()
