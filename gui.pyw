# gui.py
import os
import subprocess
import sys

import customtkinter as ctk

import main


class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.original_stdout = sys.__stdout__

    def write(self, message):
        if not message.strip():
            return

        # === Вывод в консоль (всегда безопасно) ===
        self.original_stdout.write(message)
        self.original_stdout.flush()

        # === Вывод в GUI — с защитой от ошибок ===
        try:
            if self.text_widget.winfo_exists():  # Проверяем, жив ли виджет
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", message)
                self.text_widget.see("end")
                self.text_widget.configure(state="disabled")
                self.text_widget.update_idletasks()  # Принудительное обновление
        except Exception:
            # Игнорируем ошибки GUI (например, если окно закрыто)
            pass

    def flush(self):
        self.original_stdout.flush()


# === Класс для всплывающих подсказок ===
class ToolTip:
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay  # задержка перед показом (мс)
        self.tip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<ButtonPress>", self.on_leave)

    def on_enter(self, event=None):
        self.schedule()

    def on_leave(self, event=None):
        self.unschedule()
        self.hide()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.delay, self.show)

    def unschedule(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def show(self):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = ctk.CTkToplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        self.tip_window.attributes("-topmost", True)
        label = ctk.CTkLabel(
            self.tip_window,
            text=self.text,
            fg_color="gray20",
            corner_radius=6,
            padx=8,
            pady=4
        )
        label.pack()
        # Привязка к родителю (чтобы закрывалась при закрытии основного окна)
        self.widget.winfo_toplevel().bind("<Destroy>", self.on_destroy, add="+")

    def hide(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def on_destroy(self, event=None):
        self.hide()


class DealApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Управление сделками")
        self.geometry("480x520")  # увеличил высоту под лог
        self.resizable(False, False)

        # Заголовок
        ctk.CTkLabel(self, text="ID сделки:", font=("Helvetica", 14)).pack(pady=(20, 5))
        self.deal_id_entry = ctk.CTkEntry(self, width=300, font=("Helvetica", 13))
        self.deal_id_entry.pack(pady=5)

        # Кнопки с эмодзи и подсказками
        button_data = [
            ("📥 Импортировать", "Импортировать данные по ID сделки", self.import_data),
            ("🧮 Рассчитать", "Выполнить расчёт по сделке", self.calculate),
            ("📤 Экспортировать", "Экспортировать результаты", self.export_data),
            ("📄 Сформировать 3КП", "Создать коммерческое предложение", self.generate_3kp),
            ("📋 Сформировать ТЗ", "Создать техническое задание", self.generate_tz),
            ("📂 Открыть папку", "Открыть папку с результатами", self.open_folder),
        ]

        self.buttons = []
        for i in range(0, len(button_data), 2):
            frame = ctk.CTkFrame(self, fg_color="transparent")
            frame.pack(pady=8)
            for j in range(2):
                if i + j < len(button_data):
                    text, tooltip, command = button_data[i + j]
                    btn = ctk.CTkButton(
                        frame,
                        text=text,
                        width=200,
                        height=60,
                        font=("Helvetica", 12),
                        command=command
                    )
                    btn.pack(side="left", padx=10)
                    ToolTip(btn, tooltip)
                    self.buttons.append(btn)

        # === Поле лога внизу окна ===
        ctk.CTkLabel(self, text="Лог:", font=("Helvetica", 11)).pack(anchor="w", padx=20, pady=(15, 0))
        self.log_textbox = ctk.CTkTextbox(self, height=120, wrap="word", font=("Consolas", 10))
        self.log_textbox.pack(fill="both", padx=20, pady=(5, 15), expand=False)
        self.log_textbox.configure(state="disabled")

        # Перенаправление print() в GUI
        # sys.stdout = StdoutRedirector(self.log_textbox)

    def show_message(self, title: str, message: str):
        """Показывает модальное окно с сообщением (без поля ввода)."""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("350x120")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.focus()

        # Центрируем относительно родителя
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 175
        y = self.winfo_y() + (self.winfo_height() // 2) - 60
        dialog.geometry(f"+{x}+{y}")

        ctk.CTkLabel(dialog, text=message, wraplength=300, justify="center").pack(pady=(20, 10))
        ctk.CTkButton(dialog, text="OK", width=80, command=dialog.destroy).pack(pady=(0, 15))

        dialog.bind("<Return>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        dialog.focus_set()

    def get_deal_id(self):
        deal_id = self.deal_id_entry.get().strip()
        if not deal_id:
            deal_id = self.ask_deal_id()  # спрашиваем в модальном окне
            if deal_id:
                self.deal_id_entry.delete(0, "end")
                self.deal_id_entry.insert(0, deal_id)
            else:
                return None
        return deal_id

    def ask_deal_id(self):
        """Показывает модальное окно для ввода ID сделки и возвращает его."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Введите ID сделки")
        dialog.geometry("300x120")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()  # делает окно модальным

        ctk.CTkLabel(dialog, text="ID сделки:", font=("Helvetica", 12)).pack(pady=(15, 5))
        entry = ctk.CTkEntry(dialog, width=200, font=("Helvetica", 12))
        entry.pack(pady=5)
        entry.focus()

        result = [None]

        def on_ok():
            result[0] = entry.get().strip()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="OK", command=on_ok, width=80).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", command=on_cancel, width=80).pack(side="left", padx=5)

        # 🔥 Правильный способ ожидания закрытия окна:
        dialog.wait_window()  # ← именно dialog, а не self!

        return result[0]

    def show_error(self, msg):
        self.show_message("Ошибка", msg)

    def safe_call(self, func, deal_id, success_msg):
        try:
            func(deal_id)
            self.show_message("Готово", success_msg)
        except Exception as e:
            self.show_error(f"Ошибка: {e}")

    def import_data(self):
        if deal_id := self.get_deal_id():
            self.safe_call(main.import_data, deal_id, f"Данные для {deal_id} импортированы.")

    def calculate(self):
        if deal_id := self.get_deal_id():
            self.safe_call(main.calculate, deal_id, f"Расчёт для {deal_id} завершён.")

    def export_data(self):
        if deal_id := self.get_deal_id():
            self.safe_call(main.export_data, deal_id, f"Результаты экспортированы.")

    def generate_3kp(self):
        if deal_id := self.get_deal_id():
            self.safe_call(main.generate_3kp, deal_id, f"3КП для {deal_id} создано.")

    def generate_tz(self):
        if deal_id := self.get_deal_id():
            self.safe_call(main.generate_tz, deal_id, f"ТЗ для {deal_id} создано.")

    def open_folder(self):
        deal_id = self.get_deal_id()  # ← используем общий метод
        if not deal_id:
            return  # пользователь отменил ввод

        # Путь: текущая папка / Расчеты / {deal_id}
        current_dir = os.path.dirname(os.path.abspath(__file__))
        target_dir = os.path.join(current_dir, "Расчеты", deal_id)

        os.makedirs(target_dir, exist_ok=True)

        try:
            if sys.platform == "win32":
                os.startfile(target_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", target_dir])
            else:
                subprocess.run(["xdg-open", target_dir])
        except Exception as e:
            self.show_error(f"Не удалось открыть папку:\n{e}")


if __name__ == "__main__":
    app = DealApp()
    app.mainloop()
