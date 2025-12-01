import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading
import sys
import os

# Import các hàm DLSSP
try:
    from dlssp_pipeline import run_pipeline_from_excel
except ImportError as e:
    messagebox.showerror("Import Error", f"Cannot import dlssp_pipeline:\n{e}")
    raise

class DLSSPGUI:
    def __init__(self, master):
        self.master = master
        master.title("DLSSP - Loop Sorter Scheduling")

        # File selection
        self.label = tk.Label(master, text="Select input Excel file:")
        self.label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.file_path_var = tk.StringVar()
        self.entry = tk.Entry(master, width=60, textvariable=self.file_path_var)
        self.entry.grid(row=0, column=1, padx=5, pady=5)

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        # Run button
        self.run_button = tk.Button(master, text="Run DLSSP", command=self.run_dlssp_thread)
        self.run_button.grid(row=1, column=1, padx=5, pady=5)

        # Log display
        self.log_area = scrolledtext.ScrolledText(master, width=100, height=30)
        self.log_area.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

        # Redirect stdout/stderr
        sys.stdout = self
        sys.stderr = self

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file_path_var.set(filename)

    def write(self, msg):
        self.log_area.insert(tk.END, msg)
        self.log_area.see(tk.END)

    def flush(self):
        pass

    def run_dlssp_thread(self):
        # Run DLSSP in a separate thread to keep GUI responsive
        t = threading.Thread(target=self.run_dlssp)
        t.start()

    def run_dlssp(self):
        filepath = self.file_path_var.get()
        if not filepath or not os.path.isfile(filepath):
            messagebox.showwarning("Warning", "Please select a valid Excel file!")
            return
        try:
            print(f"Running DLSSP pipeline on: {filepath}")
            out_path = run_pipeline_from_excel(filepath)
            print(f"\n Done! Results saved to: {out_path}")
            messagebox.showinfo("Success", f"DLSSP completed!\nOutput saved to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"DLSSP failed:\n{e}")
            raise

if __name__ == "__main__":
    root = tk.Tk()
    gui = DLSSPGUI(root)
    root.mainloop()
