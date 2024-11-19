import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os

# GPA calculation function
def calculate_gpa(row):
    gpa = (row['Physics'] * 0.25 +
           row['Calculus'] * 0.25 +
           row['Advanced Programming'] * 0.30 +
           row['Chemistry'] * 0.20)
    return gpa

# Excel downloading function
def load_file():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        global data
        data = pd.read_excel(file_path)
        data['GPA'] = data.apply(calculate_gpa, axis=1)
        data['Rank'] = data['GPA'].rank(ascending=False, method='min')
        messagebox.showinfo("Success", "File loaded successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading the file: {str(e)}")

# Student display function
def display_student_info():
    try:
        student_id = student_id_entry.get()
        if not student_id.isdigit():
            messagebox.showerror("Error", "Please enter a valid numeric student ID.")
            return

        student_info = data[data['ID'] == int(student_id)]
        if student_info.empty:
            messagebox.showwarning("Not Found", "Student ID not found!")
        else:
            student_info = student_info.iloc[0]
            name_label.config(text=f"Name: {student_info['Name']}")
            surname_label.config(text=f"Surname: {student_info['Surname']}")
            gpa_label.config(text=f"GPA: {student_info['GPA']:.2f}")
            rank_label.config(text=f"Rank: {int(student_info['Rank'])}")
    except Exception as e:
        messagebox.showerror("Error", f"Error displaying student info: {str(e)}")

# Clear function to be displayed items
def clear_display():
    name_label.config(text="Name: ")
    surname_label.config(text="Surname: ")
    gpa_label.config(text="GPA: ")
    rank_label.config(text="Rank: ")
    student_id_entry.delete(0, tk.END)

# Data export function
def export_data():
    try:
        file_type = file_type_combo.get()
        student_id = student_id_entry.get()
        if not student_id.isdigit():
            messagebox.showerror("Error", "Please enter a valid student ID.")
            return

        student_info = data[data['ID'] == int(student_id)].iloc[0]
        filename = f"{student_info['ID']}_{student_info['Name']}_{student_info['Surname']}.{file_type.lower()}"

        if file_type == "txt":
            with open(filename, 'w') as f:
                f.write(f"ID: {student_info['ID']}\n")
                f.write(f"Name: {student_info['Name']}\n")
                f.write(f"Surname: {student_info['Surname']}\n")
                f.write(f"GPA: {student_info['GPA']:.2f}\n")
                f.write(f"Rank: {int(student_info['Rank'])}\n")
        elif file_type == "xlsx":
            student_info.to_frame().T.to_excel(filename, index=False)

        messagebox.showinfo("Exported", f"Data exported as {filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Error exporting data: {str(e)}")

# Style settings
style = ttk.Style()
style.theme_use('clam')
style.configure('TButton', background='#9370DB', foreground='white')
style.configure('TCombobox', fieldbackground='#ADD8E6', background='#ADD8E6')

# Creating a main window
window = tk.Tk()
window.title("Student GPA and Ranking")
window.geometry("400x400")
window.configure(bg="#ADD8E6")

# UI Elements
title_label = tk.Label(window, text="Student GPA and Ranking", font=("Arial", 14), bg="#9370DB", fg="white")
title_label.pack(pady=10)

browse_button = ttk.Button(window, text="Browse", command=load_file)
browse_button.pack()

student_id_frame = tk.Frame(window, bg="#ADD8E6")
student_id_frame.pack(pady=5)
student_id_label = tk.Label(student_id_frame, text="Enter Student ID:", bg="#ADD8E6")
student_id_label.pack(side=tk.LEFT)
student_id_entry = tk.Entry(student_id_frame)
student_id_entry.pack(side=tk.LEFT)

display_button = ttk.Button(window, text="Display", command=display_student_info)
display_button.pack(pady=5)

name_label = tk.Label(window, text="Name: ", bg="#ADD8E6")
name_label.pack()
surname_label = tk.Label(window, text="Surname: ", bg="#ADD8E6")
surname_label.pack()
gpa_label = tk.Label(window, text="GPA: ", bg="#ADD8E6")
gpa_label.pack()
rank_label = tk.Label(window, text="Rank: ", bg="#ADD8E6")
rank_label.pack()

button_frame = tk.Frame(window, bg="#ADD8E6")
button_frame.pack(pady=10)

clear_button = ttk.Button(button_frame, text="Clear", command=clear_display)
clear_button.grid(row=0, column=0, padx=5)

file_type_combo = ttk.Combobox(button_frame, values=["txt", "xlsx"], state="readonly")
file_type_combo.set("txt")
file_type_combo.grid(row=0, column=1, padx=5)

export_button = ttk.Button(button_frame, text="Export", command=export_data)
export_button.grid(row=0, column=2, padx=5)


window.mainloop()
