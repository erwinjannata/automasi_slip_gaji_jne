from functions.function import generate_slip, send_email
import sys
import os
import threading
from tkinter.messagebox import showinfo
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from PIL import Image
import customtkinter
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.configure(bg="white")
root.geometry("350x335")
root.title("Automasi Slip Gaji JNE")

file_data = ''

file_data_name = tk.StringVar()

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def load_data():
    global file_data
    file_data = filedialog.askopenfilename(filetypes=(
        ("Excel Workbook", "*.xlsx"),
        ("All Files", "*.*"),
    ))
    file_data_name.set(os.path.split(file_data)[1])


def generate_progress():
    if (os.path.exists(file_data)):
        progressbar.start()
        generate_slip(file_data=file_data)
    else:
        showinfo(title='Message', message='File data tidak ditemukan!')

    btn1.configure(state="normal")
    btn2.configure(state="normal")
    btn3.configure(state="normal")


def generate_smtp():
    if (os.path.exists(file_data)):
        progressbar.start()
        send_email(file_data=file_data)
    else:
        showinfo(title='Message', message='File data tidak ditemukan!')

    btn1.configure(state="normal")
    btn2.configure(state="normal")
    btn3.configure(state="normal")


def start_thread(mode, title):
    response = messagebox.askquestion('Konfirmasi', f'Mulai Proses {title}?')

    if response == 'yes':
        global generate_thread
        if mode == 1:
            generate_thread = threading.Thread(target=generate_progress)
        else:
            generate_thread = threading.Thread(target=generate_smtp)
        generate_thread.daemon = True
        btn1.configure(state="disabled")
        btn2.configure(state="disabled")
        btn3.configure(state="disabled")
        generate_thread.start()
        root.after(20, check_thread)


def check_thread():
    if generate_thread.is_alive():
        root.after(20, check_thread)
    else:
        progressbar.stop()


# GUI
jne_image = customtkinter.CTkImage(
    light_image=Image.open(resource_path('images/jne.ico')), size=(225, 85))

image_label = customtkinter.CTkLabel(root, image=jne_image, text="").pack(
    fill="x", padx=10, pady=10)

bold_font = customtkinter.CTkFont(
    family='Arial', size=18, weight='bold', slant='italic')
label1 = customtkinter.CTkLabel(
    root, text='File Data Gaji Karyawan', anchor='w', font=bold_font).pack(fill="x", padx=10, pady=5)

label_name1 = customtkinter.CTkLabel(root, textvariable=file_data_name, anchor='w').pack(
    fill="x", padx=10, pady=5)

btn1 = customtkinter.CTkButton(
    root, text="Pilih File", command=load_data, state="normal")
btn1.pack(fill="x", padx=10, pady=5)

separator = ttk.Separator(root, orient='horizontal').pack(
    fill='x', pady=5, padx=10)

btn2 = customtkinter.CTkButton(root, text="Buat Slip Gaji",
                               command=lambda: start_thread(mode=1, title='Pembuatan Slip Gaji'), state="normal")
btn2.pack(fill="x", padx=10, pady=5)

btn3 = customtkinter.CTkButton(root, text="Kirim via Email",
                               command=lambda: start_thread(mode=2, title='Pengiriman Email'), state="normal")
btn3.pack(fill="x", padx=10, pady=5)


progressbar = customtkinter.CTkProgressBar(
    root, mode='indeterminate', orientation='horizontal')
progressbar.pack(fill='x', padx=10, pady=10)

root.mainloop()
