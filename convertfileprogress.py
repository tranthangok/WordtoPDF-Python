import tkinter as tk
from tkinter import filedialog, messagebox
from function import Converter
import os

class FileConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")
        self.root.geometry("600x500")

        self.input_path_label = tk.Label(self.root, text="Input file:")
        self.input_path_label.pack()

        self.input_path_var = tk.StringVar()
        self.input_path_entry = tk.Entry(self.root, textvariable=self.input_path_var, width=50, state="readonly")
        self.input_path_entry.pack()

        self.drop_target = tk.Label(self.root, text="Drop file or click to choose file/folder", bg="lightgray", width=50, height=5)
        self.drop_target.pack(pady=10)

        self.output_label = tk.Label(self.root, text="Output directory:")
        self.output_label.pack()

        self.output_path_var = tk.StringVar()
        self.output_path_entry = tk.Entry(self.root, textvariable=self.output_path_var, width=50, state="disabled")
        self.output_path_entry.pack()

        self.output_button = tk.Button(self.root, text="Select Output Directory", command=self.choose_output_directory)
        self.output_button.pack(pady=10)

        self.filename_label = tk.Label(self.root, text="Output filename:")
        self.filename_label.pack()

        self.filename_entry = tk.Entry(self.root, width=50)
        self.filename_entry.pack()

        self.convert_button = tk.Button(self.root, text="Convert", command=self.convert_file)
        self.convert_button.pack(pady=20)

        self.input_path = None
        self.output_path = None

        self.drop_target.bind("<Enter>", self.on_enter)
        self.drop_target.bind("<Leave>", self.on_leave)
        self.drop_target.bind("<Button-1>", self.on_click_drop_target)

    def choose_output_directory(self):
        self.output_path = filedialog.askdirectory()
        if self.output_path:
            self.output_path_var.set(self.output_path)
            self.output_label.config(text=f"Output directory:")
    
    def convert_file(self):
        if self.input_path and self.output_path:
            output_filename = self.filename_entry.get()
            if not output_filename:
                messagebox.showwarning("Warning", "Please enter an output filename.")
                return

            output_path = os.path.join(self.output_path, output_filename + ".pdf")

            try:
                if self.input_path.endswith(".docx") or self.input_path.endswith(".doc"):
                    self.convert_with_progress(Converter.convert_docx_to_pdf, output_path)
                elif self.input_path.endswith(".xlsx") or self.input_path.endswith(".xls"):
                    self.convert_with_progress(Converter.convert_xlsx_to_pdf, output_path)
                elif self.input_path.endswith(".pptx") or self.input_path.endswith(".ppt"):
                    self.convert_with_progress(Converter.convert_pptx_to_pdf, output_path)
                else:
                    messagebox.showwarning("Warning", "Unsupported file format.")
            except Exception as e:
                messagebox.showerror("Error", f"File conversion error: {e}")
        else:
            messagebox.showwarning("Warning", "Please select both input file and output directory.")
    
    def convert_with_progress(self, conversion_function, output_path):
        if self.input_path and self.output_path:
            self.progress_window = tk.Toplevel(self.root)
            self.progress_window.title("Converting...")
            self.progress_label = tk.Label(self.progress_window, text="Converting...")
            self.progress_label.pack()

            def update_progress(progress):
                self.progress_label.config(text=f"Converting... {progress}%")

            conversion_function(self.input_path, output_path, update_progress)
            self.progress_label.config(text="Conversion complete!")
        else:
            messagebox.showwarning("Warning", "Please select both input file and output directory.")

    def on_enter(self, event):
        self.drop_target.config(bg="lightblue")
    
    def on_leave(self, event):
        self.drop_target.config(bg="lightgray")
    
    def on_click_drop_target(self, event):
        self.input_path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
        if self.input_path:
            self.input_path_var.set(self.input_path)


