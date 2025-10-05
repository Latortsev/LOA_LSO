# gui.py
import tkinter as tk
from tkinter import messagebox
import main  # импортируем ваш основной модуль


class DealApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Управление сделками")
        self.root.geometry("400x250")

        # Метка и поле ввода ID сделки
        tk.Label(root, text="ID сделки:", font=("Arial", 12)).pack(pady=(20, 5))
        self.deal_id_entry = tk.Entry(root, font=("Arial", 12), width=30)
        self.deal_id_entry.pack(pady=5)

        # Кнопки
        button_frame = tk.Frame(root)
        button_frame.pack(pady=20)

        buttons = [
            ("Импортировать", self.import_data),
            ("Рассчитать", self.calculate),
            ("Экспортировать", self.export_data),
            ("Сформировать 3КП", self.generate_3kp),
            ("Сформировать ТЗ", self.generate_tz),
        ]

        for text, command in buttons:
            tk.Button(button_frame, text=text, command=command, width=20, height=1).pack(pady=5)

    def get_deal_id(self):
        deal_id = self.deal_id_entry.get().strip()
        if not deal_id:
            messagebox.showwarning("Ошибка", "Пожалуйста, введите ID сделки.")
            return None
        return deal_id

    def import_data(self):
        deal_id = self.get_deal_id()
        if deal_id:
            try:
                main.import_data(deal_id)
                messagebox.showinfo("Успех", f"Данные для сделки {deal_id} импортированы.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при импорте: {e}")

    def calculate(self):
        deal_id = self.get_deal_id()
        if deal_id:
            try:
                main.calculate(deal_id)
                messagebox.showinfo("Успех", f"Расчёт для сделки {deal_id} выполнен.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при расчёте: {e}")

    def export_data(self):
        deal_id = self.get_deal_id()
        if deal_id:
            try:
                main.export_data(deal_id)
                messagebox.showinfo("Успех", f"Данные для сделки {deal_id} экспортированы.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")

    def generate_3kp(self):
        deal_id = self.get_deal_id()
        if deal_id:
            try:
                main.generate_3kp(deal_id)
                messagebox.showinfo("Успех", f"3КП для сделки {deal_id} сформировано.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при формировании 3КП: {e}")

    def generate_tz(self):
        deal_id = self.get_deal_id()
        if deal_id:
            try:
                main.generate_tz(deal_id)
                messagebox.showinfo("Успех", f"ТЗ для сделки {deal_id} сформировано.")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при формировании ТЗ: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DealApp(root)
    root.mainloop()