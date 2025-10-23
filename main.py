# EDUMETRICS main (simplified starter)
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd, os
SCHOOL_NAME = "ELGEYO SAWMILL PRIMARY AND JUNIOR SCHOOL"
def run_app():
    root = tk.Tk()
    root.title("EDUMETRICS - Exam Analysis System")
    root.state("zoomed")
    tk.Label(root, text=SCHOOL_NAME, font=("Segoe UI",20)).pack(pady=10)
    def open_sample():
        p = os.path.join(os.getcwd(),"SampleResults.xlsx")
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showerror("Missing", "SampleResults.xlsx not found")
    tk.Button(root, text="Open SampleData.xlsx", command=open_sample).pack(pady=6)
    root.mainloop()
if __name__=='__main__':
    run_app()
