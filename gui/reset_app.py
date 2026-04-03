# gui/reset_app.py
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import os
from datetime import datetime

from reset_computed_cells import (
    reset_computed_kot1_cells, reset_computed_kot2_cells,
    reset_computed_ppk_cells, reset_computed_pwk_cells,
    reset_calc_tep_cells, reset_computed_tur1_cells,
    reset_computed_tur2_cells, reset_computed_zagal_cells
)
from .base_tab import BaseTab


class ResetApp(BaseTab):
    def __init__(self, parent_frame):
        super().__init__(parent_frame)

        tk.Label(self.frame, text="Обнулити розрахункові клітинки",
                 font=("Arial", 16, "bold"), fg="#d32f2f").pack(pady=15)

        # Вибір файлу
        file_frame = tk.Frame(self.frame)
        file_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(file_frame, text="Excel файл:", font=("Arial", 11)).pack(side="left")
        tk.Entry(file_frame, textvariable=self.file_path, width=55, state="readonly").pack(side="left", padx=10)
        tk.Button(file_frame, text="Обрати файл", command=self.select_file, width=15).pack(side="left")

        # Вибір сторінок
        tk.Label(self.frame, text="Оберіть сторінки, які потрібно обнулити:",
                 font=("Arial", 11, "bold")).pack(anchor="w", padx=20, pady=(20, 5))

        self.select_all_var = tk.IntVar(value=1)
        self.select_all_cb = tk.Checkbutton(
            self.frame,
            text="✅ Вибрати ВСІ сторінки",
            variable=self.select_all_var,
            command=self.toggle_select_all,
            font=("Arial", 10)
        )
        self.select_all_cb.pack(anchor="w", padx=40)

        self.sheet_options = [
            {"name": "Турбіна I черга", "sheet_name": "Турбіна I черга", "func": reset_computed_tur1_cells},
            {"name": "Турбіна II черга", "sheet_name": "Турбіна II черга", "func": reset_computed_tur2_cells},
            {"name": "Котел I черга", "sheet_name": "Котел I черга", "func": reset_computed_kot1_cells},
            {"name": "Котел II черга", "sheet_name": "Котел II черга", "func": reset_computed_kot2_cells},
            {"name": "ПВК", "sheet_name": "ПВК", "func": reset_computed_pwk_cells},
            {"name": "ТЕП", "sheet_name": "ТЕП", "func": reset_calc_tep_cells},
            {"name": "ППК", "sheet_name": "ППК", "func": reset_computed_ppk_cells},
            {"name": "Загальні", "sheet_name": "Загальні", "func": reset_computed_zagal_cells}
        ]

        self.reset_vars = {}
        for opt in self.sheet_options:
            var = tk.IntVar(value=1)
            cb = tk.Checkbutton(
                self.frame,
                text=opt["name"],
                variable=var,
                font=("Arial", 10)
            )
            cb.pack(anchor="w", padx=40)
            self.reset_vars[opt["name"]] = var

        self.reset_btn = tk.Button(
            self.frame,
            text="ОБНУЛИТИ ВИБРАНІ СТОРІНКИ",
            font=("Arial", 12, "bold"),
            bg="#d32f2f",
            fg="white",
            height=2,
            command=self.run_reset
        )
        self.reset_btn.pack(pady=25)

        self.status = tk.Label(self.frame, text="Готовий", fg="green", font=("Arial", 10))
        self.status.pack(pady=5)

        self.create_log_widget(height=12)

    def toggle_select_all(self):
        state = self.select_all_var.get()
        for var in self.reset_vars.values():
            var.set(state)

    def run_reset(self):
        if not self.file_path.get():
            messagebox.showerror("Помилка", "Спочатку оберіть Excel-файл!")
            return

        selected = [opt for opt in self.sheet_options if self.reset_vars[opt["name"]].get() == 1]

        if not selected:
            messagebox.showwarning("Попередження", "Не вибрано жодної сторінки для обнуління!")
            return

        self.reset_btn.config(state="disabled")
        self.status.config(text="Обнулення в процесі...", fg="orange")
        self.log("🚀 Початок обнуління...")

        try:
            wb = openpyxl.load_workbook(self.file_path.get(), data_only=False)

            for opt in selected:
                try:
                    sheet = wb[opt["sheet_name"]]
                    opt["func"](sheet)
                    self.log(f"✅ Обнулено: {opt['name']}")
                except KeyError:
                    self.log(f"⚠️ Аркуш «{opt['sheet_name']}» не знайдено в файлі!")
                except Exception as e:
                    self.log(f"❌ Помилка в {opt['name']}: {e}")

            base, ext = os.path.splitext(self.file_path.get())
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_path = f"{base}_ОБНУЛЕНО_{timestamp}{ext}"
            wb.save(new_path)

            self.log("🎉 Обнулення завершено успішно!")
            self.log(f"Файл збережено: {os.path.basename(new_path)}")
            self.status.config(text="Успішно обнулено!", fg="green")

            messagebox.showinfo(
                "Готово!",
                f"Обнулено {len(selected)} сторінок\n\nЗбережено як:\n{os.path.basename(new_path)}"
            )

        except Exception as e:
            self.log(f"🔥 Критична помилка: {e}")
            messagebox.showerror("Помилка", f"Не вдалося виконати обнуління:\n{str(e)}")
            self.status.config(text="Помилка!", fg="red")
        finally:
            self.reset_btn.config(state="normal")