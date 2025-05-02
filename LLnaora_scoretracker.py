import tkinter as tk
from tkinter import *
from openpyxl import Workbook


wb = Workbook()
sheet = wb.active
sheet.title = "Student name and score"


sheet['A1'] = "Surname"
sheet['B1'] = "Score"
sheet['C1'] = "Status"



row = 2 


def check():
    
   
    global row   
    surname = entry1.get()
    grades = entry2.get()

    
    
    grade = int(grades)
    if grade >= 1 and grade <= 74:
            status = "Failed"
    elif grade >= 75 and grade <= 100:
            status = "Passed"
    else:
            status = "Invalid Score"
    
    

   
    sheet[f'A{row}'] = surname
    sheet[f'B{row}'] = grade
    sheet[f'C{row}'] = status
    row += 1

 
            
    wb.save("student_score_1.xlsx")

   
    entry1.delete(0, END)
    entry2.delete(0, END)

window = tk.Tk()
window.geometry("300x200")
window.title("Grade Checker")

sn = tk.Label(window, text="Surname")
sn.pack()
entry1 = tk.Entry(window)
entry1.pack()

grd = tk.Label(window, text="Grade")
grd.pack()
entry2 = tk.Entry(window)
entry2.pack()

cs = tk.Button(window, text="Check and Save", command=check)
cs.pack(pady=10)

window.mainloop()