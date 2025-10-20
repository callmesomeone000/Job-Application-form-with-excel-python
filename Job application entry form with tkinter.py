from tkinter import *
from tkinter.ttk import Combobox
import openpyxl
from openpyxl import Workbook
import pathlib
from tkinter import messagebox

# ---------------- File setup ----------------
file = pathlib.Path('Job_Applications.xlsx')
if not file.exists():
    wb = Workbook()
    sheet = wb.active
    sheet.append(["Full Name", "Contact Number", "Age", "Gender", "Address", "Qualification", "Experience (Years)", "Skills"])
    wb.save('Job_Applications.xlsx')

# ---------------- Functions ----------------
def submit():
    name = name_var.get()
    contact = contact_var.get()
    age = age_var.get()
    gender = gender_box.get()
    address = address_entry.get(1.0, END).strip()
    qualification = qualification_var.get()
    experience = experience_var.get()
    skills = skills_entry.get(1.0, END).strip()

    if not name or not contact:
        messagebox.showwarning("Warning", "Please fill in all required fields.")
        return

    wb = openpyxl.load_workbook('Job_Applications.xlsx')
    sheet = wb.active
    sheet.append([name, contact, age, gender, address, qualification, experience, skills])
    wb.save('Job_Applications.xlsx')

    messagebox.showinfo("Success", "Details added successfully!")
    clear_fields()

def clear_fields():
    name_var.set('')
    contact_var.set('')
    age_var.set('')
    qualification_var.set('')
    experience_var.set('')
    address_entry.delete(1.0, END)
    skills_entry.delete(1.0, END)
    gender_box.set('Male')

# ---------------- UI Setup ----------------
root = Tk()
root.title("Job Application Form")
root.geometry("850x600+300+100")
root.config(bg="#2B2B2B")  # soft black background

# Title
title = Label(root, text="📝 Job Application Entry Form", font=("Segoe UI Semibold", 20), bg="#2B2B2B", fg="white")
title.pack(pady=20)

# Frame
form_frame = Frame(root, bg="#2B2B2B")
form_frame.pack(pady=10)

# Variables
name_var = StringVar()
contact_var = StringVar()
age_var = StringVar()
qualification_var = StringVar()
experience_var = StringVar()

# Full Name
Label(form_frame, text="Full Name", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=0, column=0, padx=20, pady=10, sticky=W)
Entry(form_frame, textvariable=name_var, width=50, font=("Segoe UI", 10), bd=1, bg="white", fg="black").grid(row=0, column=1, pady=5, sticky=W)

# Contact
Label(form_frame, text="Contact Number", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=1, column=0, padx=20, pady=10, sticky=W)
Entry(form_frame, textvariable=contact_var, width=50, font=("Segoe UI", 10), bd=1, bg="white", fg="black").grid(row=1, column=1, pady=5, sticky=W)

# Age and Gender
Label(form_frame, text="Age", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=2, column=0, padx=20, pady=10, sticky=W)
age_entry = Entry(form_frame, textvariable=age_var, width=15, font=("Segoe UI", 10), bd=1, bg="white", fg="black")
age_entry.grid(row=2, column=1, pady=5, sticky=W)

Label(form_frame, text="Gender", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=2, column=1, padx=150, pady=10, sticky=W)
gender_box = Combobox(form_frame, values=["Male", "Female", "Other"], font=("Segoe UI", 10), width=15, state="readonly")
gender_box.set("Male")
gender_box.grid(row=2, column=1, padx=220, pady=5, sticky=W)

# Address
Label(form_frame, text="Address", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=3, column=0, padx=20, pady=10, sticky=W)
address_entry = Text(form_frame, width=50, height=4, font=("Segoe UI", 10), bd=1, bg="white", fg="black")
address_entry.grid(row=3, column=1, pady=5, sticky=W)

# Qualification
Label(form_frame, text="Qualification", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=4, column=0, padx=20, pady=10, sticky=W)
Entry(form_frame, textvariable=qualification_var, width=50, font=("Segoe UI", 10), bd=1, bg="white", fg="black").grid(row=4, column=1, pady=5, sticky=W)

# Experience
Label(form_frame, text="Experience (Years)", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=5, column=0, padx=20, pady=10, sticky=W)
Entry(form_frame, textvariable=experience_var, width=50, font=("Segoe UI", 10), bd=1, bg="white", fg="black").grid(row=5, column=1, pady=5, sticky=W)

# Skills
Label(form_frame, text="Skills", font=("Segoe UI", 11, "bold"), bg="#2B2B2B", fg="white").grid(row=6, column=0, padx=20, pady=10, sticky=W)
skills_entry = Text(form_frame, width=50, height=4, font=("Segoe UI", 10), bd=1, bg="white", fg="black")
skills_entry.grid(row=6, column=1, pady=5, sticky=W)

# ---------------- Buttons ----------------
def on_enter(e): 
    e.widget.config(bg="#3A3A3A")

def on_leave(e): 
    e.widget.config(bg="#2B2B2B")

button_frame = Frame(root, bg="#2B2B2B")
button_frame.pack(pady=30)

# Button style with border
btn_style = {
    "font": ("Segoe UI Semibold", 11),
    "fg": "white",
    "bg": "#2B2B2B",
    "width": 15,
    "height": 2,
    "bd": 2,             # border thickness
    "relief": "ridge",   # 3D border style
    "activebackground": "#3A3A3A"
}

submit_btn = Button(button_frame, text="Submit", command=submit, **btn_style)
clear_btn = Button(button_frame, text="Clear", command=clear_fields, **btn_style)
exit_btn = Button(button_frame, text="Exit", command=root.destroy, **btn_style)

for btn in (submit_btn, clear_btn, exit_btn):
    btn.pack(side=LEFT, padx=20)
    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)

root.mainloop()
