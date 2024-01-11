import tkinter as tk
from tkinter import filedialog
import os


root = tk.Tk()
root.geometry("550x550")
root.withdraw()
root_dir = os.path.dirname(os.path.abspath(__file__))

file_path = filedialog.askopenfilename(initialdir = root_dir+'/template/', title = "Select a DOCX file", filetypes = (("DOCX files", "*.docx"), ("all files", "*.*")))

print(file_path)