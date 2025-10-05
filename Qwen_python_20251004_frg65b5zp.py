# gui.py
import os
import customtkinter as ctk
from PIL import Image
import main  # ваш модуль с функциями
import subprocess
import sys


# Настройка внешнего вида
ctk.set_appearance_mode("dark")  # "light" или "dark"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"


class DealApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Управление сделками")
        self.geometry("500x400")
        self.resizable(False, False)

        # Путь к иконкам
        self.icon_path = os.path.join(os.path.dirname(__file__), "icons")

        # Загрузка иконок
        self.icons = {
            "import": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "import.png")), size=(24, 24)),
            "calculate": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "calculate.png")), size=(24, 24)),
            "export": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "export.png")), size=(24, 24)),
            "kp3": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "kp3.png")), size=(24, 24)),
            "tz": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "tz.png")), size=(24, 24)),
            "folder": ctk.CTkImage(Image.open(os.path.join(self.icon_path, "folder.png")), size=(24, 24)),
        }

        # Поле ввода ID сделки
        self.deal_id_label = ctk.CTkLabel(self, text="ID сделки:", font=("Helvetica", 14))
        self.deal_id_label.pack(pady=(20, 5))

        self.deal_id_entry = ctk.CTkEntry(self, width=300, font=("Helvetica", 13))
        self.deal_id_entry.pack(pady=5)

        # Контейнер для кнопок
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.pack(pady=20)

        # Создание кнопок с иконками
        buttons = [
            ("Импортировать", self.icons["import"], self.import_data),
            ("Рассчитать", self.icons["calculate"], self.calculate),
            ("Экспортировать", self.icons["export"], self.export_data),
            ("Сформировать 3КП", self.icons["kp3"], self.generate_3kp),
            ("Сформировать ТЗ", self.icons["tz"], self.generate_tz),
            ("Открыть папку", self.icons["folder"], self.open_folder),
        ]

        # Размещаем кнопки в сетке 2×3
        for i, (text, icon, command) in enumerate(buttons):
            row = i // 2
            col = i % 2
            btn = ctk.CTkButton(
                self.button_frame,
                text=text,
                image=icon,
                compound="top",  # иконка сверху, текст снизу
                width=180,
                height=80,
                font=("Helvetica", 11),
                command=command
            )
            btn.grid(row=row, column=col, padx=10, pady=10)

    def get_deal_id(self):
        deal_id = self.deal_id_entry.get().strip()
        if not deal_id:
            self.show_error("Пожалуйста, введите ID сделки.")
            return None
        return deal_id

    def show_error(self, message):
        # Используем встроенное окно
        ctk.CTkInputDialog(text=message, title="Ошибка")  # или messagebox
        # Альтернатива: tkinter.messagebox.showerror("Ошибка", message)

    def safe_call(self, func, deal_id, success_msg):
        try:
            func(deal_id)
            ctk.CTkInputDialog(text=success_msg, title="Успех")
        except Exception as e:
            self.show_error(f"Ошибка: {str(e)}")

    # Обработчики кнопок
    def import_data(self):
        if (deal_id := self.get_deal_id()):
            self.safe_call(main.import_data, deal_id, f"Данные для сделки {deal_id} импортированы.")

    def calculate(self):
        if (deal_id := self.get_deal_id()):
            self.safe_call(main.calculate, deal_id, f"Расчёт для сделки {deal_id} завершён.")

    def export_data(self):
        if (deal_id := self.get_deal_id()):
            self.safe_call(main.export_data, deal_id, f"Данные экспортированы для сделки {deal_id}.")

    def generate_3kp(self):
        if (deal_id := self.get_deal_id()):
            self.safe_call(main.generate_3kp, deal_id, f"3КП для сделки {deal_id} сформировано.")

    def generate_tz(self):
        if (deal_id := self.get_deal_id()):
            self.safe_call(main.generate_tz, deal_id, f"ТЗ для сделки {deal_id} сформировано.")

    def open_folder(self):
        """Открывает папку с результатами (например, ./output/)"""
        output_dir = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(output_dir, exist_ok=True)

        if sys.platform == "win32":
            os.startfile(output_dir)
        elif sys.platform == "darwin":  # macOS
            subprocess.Popen(["open", output_dir])
        else:  # Linux
            subprocess.Popen(["xdg-open", output_dir])


if __name__ == "__main__":
    app = DealApp()
    app.mainloop()