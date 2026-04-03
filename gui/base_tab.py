# gui/base_tab.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

class BaseTab:
    def __init__(self, parent_frame):
        self.frame = parent_frame
        self.file_path = tk.StringVar()  # використовується не в усіх вкладках
        self.log_text = None

    def create_log_widget(self, height=10, width=85):
        self.log_text = tk.Text(self.frame, height=height, width=width, state="disabled")
        self.log_text.pack(pady=10, padx=20)

    def log(self, message: str):
        if not self.log_text:
            return
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def select_file(self, title="Оберіть Excel файл", filetypes=None):
        if filetypes is None:
            filetypes = [("Excel файли", "*.xlsx *.xlsm")]
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            self.file_path.set(path)
            self.log(f"Обрано: {os.path.basename(path)}")
            return path
        return None