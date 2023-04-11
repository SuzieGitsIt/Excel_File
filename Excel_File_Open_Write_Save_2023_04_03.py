# https://stackoverflow.com/questions/62791280/transferring-entry-fields-to-excel-file-using-pythons-tkinter

from openpyxl import *
from functools import partial                           # for allowing 015 040 buttons to equal specific values when clicked.
import datetime as dt                                   # Date library
import tkinter as tk
import tkinter.ttk as ttk

date = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

GUI = tk.Tk()
GUI.title('LAL Measurements')
GUI.geometry("600x400")

def close_window():
    GUI.destroy()

try:
    workbook = load_workbook(filename="WO+Date+Meas.xlsx")
    sheet = workbook.active
except:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Credentials"
    sheet["B1"] = "Work Order"
    sheet["C1"] = "Sample Size"
    sheet["D1"] = "Measurement Size"
    sheet["E1"] = "Date and Time"

new_line = sheet.max_row + 1

def excel_entry():
    sheet.cell(column=1, row=new_line, value=entry_cred.get()[:3])
    sheet.cell(column=2, row=new_line, value=entry_WO.get()[:10])
    sheet.cell(column=3, row=new_line, value=entry_samp.get()[:2])
    sheet.cell(column=5, row=new_line).value = date
    workbook.save(filename="WO+Date+Meas.xlsx")

def entry_015(text):
    sheet.cell(column=4, row=new_line).value = text
    excel_entry()

def entry_040(text):
    sheet.cell(column=4, row=new_line).value = text
    excel_entry()

def entry_100(text):
    sheet.cell(column=4, row=new_line).value = text
    excel_entry()

btn_exit = tk.Button(GUI, text="Exit", fg='black', command= close_window)
btn_exit.place(x=400, y=250)

lbl_disp_cred = tk.Label(GUI, text="Credentials: ", fg='black', font=("Helvetica", 8))
lbl_disp_cred.place(x=20, y=50)

lbl_disp_WO = tk.Label(GUI, text="Work Order: ", fg='black', font=("Helvetica", 8))
lbl_disp_WO.place(x=20, y=75)

lbl_disp_samp = tk.Label(GUI, text="Sample Size: ", fg='black', font=("Helvetica", 8))
lbl_disp_samp.place(x=20, y=100)

lbl_disp_meas = tk.Label(GUI, text="Measurement Size: ", fg='black', font=("Helvetica", 8))
lbl_disp_meas.place(x=20, y=125)

entry_cred = tk.Entry(GUI, bg='white',fg='black', bd=5)
entry_cred.place(x=350, y=50)

entry_WO = tk.Entry(GUI, bg='white', fg='black', bd=5)
entry_WO.place(x=350, y=75)

entry_samp = tk.Entry(GUI, bg='white', fg='black', bd=5)
entry_samp.place(x=350, y=100)

btn_015 = ttk.Button(GUI, text= "Posterior", command= partial(entry_015, "Post - 015"))
btn_015.place(x=325, y=150)
btn_040 = ttk.Button(GUI, text= "Anterior", command= partial(entry_040, "Ant - 040"))
btn_040.place(x=400, y=150)
btn_100 = ttk.Button(GUI, text= "Full Lens", command= partial(entry_100, "Full - 100"))
btn_100.place(x=475, y=150)

btn_sav = ttk.Button(GUI, text="Save Entry", command= excel_entry)
btn_sav.place(x=300, y=250)

GUI.mainloop()
