# main.py
import tkinter as tk
from tkinter import ttk
import os

from gui.import_dat_app import ImportDatApp
from gui.kot_calculator_app import KotCalculatorApp
from gui.reset_app import ResetApp

if __name__ == "__main__":
    root = tk.Tk()
    root.title("TEC — Повний інструмент (Калькулятор + Імпорт + Обнулити)")
    root.geometry("920x680")
    root.resizable(True, True)

    style = ttk.Style()
    style.theme_use('clam')

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # Вкладки
    tab_calc = ttk.Frame(notebook)
    notebook.add(tab_calc, text="Калькулятор котлів")
    KotCalculatorApp(tab_calc)

    tab_import = ttk.Frame(notebook)
    notebook.add(tab_import, text="Імпорт .dat → Excel")
    ImportDatApp(tab_import)

    tab_reset = ttk.Frame(notebook)
    notebook.add(tab_reset, text="Обнулити клітинки")
    ResetApp(tab_reset)

    root.mainloop()