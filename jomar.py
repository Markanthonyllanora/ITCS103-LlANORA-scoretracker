import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook, Workbook
import os

# Gumawa ng Excel file kung wala pa
if not os.path.exists("Grade.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Grades"
    ws.append(["Name", "Course", "Grade"])
    wb.save("Grade.xlsx")

def validate_inputs():
    name = name_entry.get()
    course = course_entry.get()
    grade = grade_entry.get()

    if not name or not course or not grade:
        messagebox.showerror("Input Error", "All fields are required!")
        return False
    return True

def save_to_excel():
    if not validate_inputs():
        return

    name = name_entry.get()
    course = course_entry.get()
    grade = grade_entry.get()

    wb = load_workbook("Grade.xlsx")
    ws = wb["Grades"]
    ws.append([name, course, grade])
    wb.save("Grade.xlsx")
    messagebox.showinfo("Success", "Data saved successfully!")

    name_entry.delete(0, tk.END)
    course_entry.delete(0, tk.END)
    grade_entry.delete(0, tk.END)

def show_data():
    wb = load_workbook("Grade.xlsx")
    ws = wb["Grades"]

    data = tk.Toplevel(window)
    data.title("Student Data")
    data.geometry("300x300")
    data.configure(bg="light blue")

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        for j, value in enumerate(row):
            label = tk.Label(data, text=value, bg="light blue", padx=6, pady=3)
            label.grid(row=i, column=j)

window = tk.Tk()
window.geometry("300x300")
window.title("Grade Report")
window.configure(bg="light blue")

header_lbl = tk.Label(window, text="Grade Report", font=("arial", 18), bg="light blue")
frame = tk.Frame(window, bg="light blue")

name_lbl = tk.Label(frame, text="Name", font=("arial", 12), bg="light blue")
course_lbl = tk.Label(frame, text="Course", font=("arial", 12), bg="light blue")
grade_lbl = tk.Label(frame, text="Grade", font=("arial", 12), bg="light blue")

name_entry = tk.Entry(frame, width=20)
course_entry = tk.Entry(frame, width=20)
grade_entry = tk.Entry(frame, width=20)

header_lbl.pack(pady=25)
frame.pack()

name_lbl.grid(row=0, column=0, sticky="w")
course_lbl.grid(row=1, column=0, sticky="w")
grade_lbl.grid(row=2, column=0, sticky="w")

name_entry.grid(row=0, column=1, padx=5, pady=5)
course_entry.grid(row=1, column=1, padx=5, pady=5)
grade_entry.grid(row=2, column=1, padx=5, pady=5)

save_btn = tk.Button(frame, text="Save", font=("arial", 9), width=12, command=save_to_excel)
view_data_btn = tk.Button(frame, text="View Data", font=("arial", 9), width=12, command=show_data)

save_btn.grid(row=3, column=0, padx=5, pady=15)
view_data_btn.grid(row=3, column=1, padx=5, pady=15)

window.mainloop()