# gui.py
import os
import subprocess
import sys
import customtkinter as ctk


# === Перенаправление stdout в текстовое поле GUI ===
class StdoutRedirector:
    """
    Перехватывает вывод print() и дублирует его в текстовое поле GUI и в консоль.
    Это удобно для логирования в приложении и для отладки из консоли.
    """
    def __init__(self, text_widget):
        self.text_widget = text_widget
        # Сохраняем оригинальный stdout; sys.__stdout__ безопаснее, если кто-то уже переназначал sys.stdout
        self.original_stdout = sys.__stdout__ or sys.stdout

    def write(self, message):
        if not message.strip():
            return

        # Печать в консоль (если доступна)
        try:
            if self.original_stdout:
                self.original_stdout.write(message)
                self.original_stdout.flush()
        except Exception:
            # Не роняем GUI, если консоль недоступна
            pass

        # Печать в GUI
        try:
            if self.text_widget.winfo_exists():
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", message)
                self.text_widget.see("end")  # автопрокрутка
                self.text_widget.configure(state="disabled")
        except Exception:
            pass

    def flush(self):
        """Совместимость с интерфейсом файла."""
        try:
            if self.original_stdout:
                self.original_stdout.flush()
        except Exception:
            pass


# === Всплывающие подсказки, которые не «залипают» ===
class ToolTip:
    """
    Надёжная подсказка для виджета.
    Особенности:
    - Задержка перед показом.
    - Автоматическое скрытие при уходе курсора с виджета или самой подсказки.
    - Дополнительный «сторож» (watchdog): периодически проверяет положение курсора
      и закрывает подсказку, если курсор ушёл (устраняет «висящие» подсказки).
    """
    def __init__(self, widget, text, delay=500, watch_interval=150):
        """
        widget: виджет, к которому привязана подсказка
        text: текст подсказки
        delay: задержка перед показом, мс
        watch_interval: период проверки положения курсора, мс
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.watch_interval = watch_interval

        self.tip_window = None
        self.show_id = None        # id after() для запланированного показа
        self.watch_id = None       # id after() для сторожа

        # Привязка событий к исходному виджету
        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")
        # На закрытие окна — закрываем подсказку
        self.widget.winfo_toplevel().bind("<Destroy>", self._on_destroy, add="+")

    # --- События виджета ---
    def _on_enter(self, _event=None):
        # Планируем показ через delay
        self._schedule_show()

    def _on_leave(self, _event=None):
        # Отменяем показ и скрываем, если уже видно
        self._unschedule_show()
        self.hide()

    # --- Планирование показа ---
    def _schedule_show(self):
        self._unschedule_show()
        self.show_id = self.widget.after(self.delay, self.show)

    def _unschedule_show(self):
        if self.show_id is not None:
            try:
                self.widget.after_cancel(self.show_id)
            except Exception:
                pass
            self.show_id = None

    # --- Показ подсказки ---
    def show(self):
        if self.tip_window or not self.text:
            return

        # Позиционируем подсказку около виджета
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        except Exception:
            return

        # Создаём всплывающее окно без рамки, поверх остальных
        self.tip_window = ctk.CTkToplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        self.tip_window.attributes("-topmost", True)

        # Контент подсказки
        label = ctk.CTkLabel(
            self.tip_window,
            text=self.text,
            fg_color="gray20",
            corner_radius=6,
            padx=8,
            pady=4,
            justify="left",
        )
        label.pack()

        # События для подсказки: уход курсора — скрыть
        self.tip_window.bind("<Leave>", lambda _e: self.hide(), add="+")
        self.tip_window.bind("<ButtonPress>", lambda _e: self.hide(), add="+")
        self.tip_window.bind("<Enter>", lambda _e: self._start_watchdog(), add="+")

        # Сразу запускаем «сторож» — на случай быстрых движений курсора
        self._start_watchdog()

    # --- Сторож: проверка положения курсора ---
    def _start_watchdog(self):
        self._stop_watchdog()
        self.watch_id = self.widget.after(self.watch_interval, self._watch_cursor)

    def _stop_watchdog(self):
        if self.watch_id is not None:
            try:
                self.widget.after_cancel(self.watch_id)
            except Exception:
                pass
            self.watch_id = None

    def _watch_cursor(self):
        """
        Проверяем, находится ли курсор над исходным виджетом или над подсказкой.
        Если нет — закрываем подсказку.
        """
        try:
            # Текущие координаты курсора на экране
            px = self.widget.winfo_pointerx()
            py = self.widget.winfo_pointery()

            def inside(win):
                if not win:
                    return False
                # Геометрия окна
                x = win.winfo_rootx()
                y = win.winfo_rooty()
                w = win.winfo_width()
                h = win.winfo_height()
                return (x <= px <= x + w) and (y <= py <= y + h)

            over_widget = inside(self.widget)
            over_tip = inside(self.tip_window)

            if not (over_widget or over_tip):
                # Курсор ушёл — скрываем и останавливаем сторож
                self.hide()
                self._stop_watchdog()
                return
        except Exception:
            # В случае ошибки просто скрываем
            self.hide()
            self._stop_watchdog()
            return

        # Продолжаем наблюдение
        self._start_watchdog()

    # --- Скрытие и уничтожение ---
    def hide(self):
        """Закрывает окно подсказки, если оно существует."""
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None

    def _on_destroy(self, _event=None):
        """При уничтожении главного окна — убираем подсказку и отменяем таймеры."""
        self._unschedule_show()
        self._stop_watchdog()
        self.hide()


# === Основное приложение ===
class DealApp(ctk.CTk):
    """
    Небольшое GUI-приложение:
    - Ввод ID сделки.
    - Кнопки для действий (импорт, расчёт, экспорт, генерация документов, открытие папки).
    - Лог внизу окна, куда выводится print().
    """
    def __init__(self):
        super().__init__()
        self.title("Расчет ЛШО")
        self.geometry("480x540")  # немного больше высота под лог и подсказки
        self.resizable(False, False)

        # — Поле ввода ID сделки —
        ctk.CTkLabel(self, text="ID сделки:", font=("Helvetica", 14)).pack(pady=(20, 5))
        self.deal_id_entry = ctk.CTkEntry(self, width=300, font=("Helvetica", 13))
        self.deal_id_entry.pack(pady=5)

        # — Кнопки действий —
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
                    ToolTip(btn, tooltip, delay=450, watch_interval=120)  # подсказки без «залипания»
                    self.buttons.append(btn)

        # — Поле лога —
        ctk.CTkLabel(self, text="Лог:", font=("Helvetica", 11)).pack(anchor="w", padx=20, pady=(15, 0))
        self.log_textbox = ctk.CTkTextbox(self, height=140, wrap="word", font=("Consolas", 10))
        self.log_textbox.pack(fill="both", padx=20, pady=(5, 15), expand=False)
        self.log_textbox.configure(state="disabled")

        # Перенаправляем stdout в лог
        sys.stdout = StdoutRedirector(self.log_textbox)

        # Импортируем main после перенаправления stdout
        import main
        self.main_module = main

    # === Утилиты для диалогов и ввода ===
    def show_message(self, title: str, message: str):
        """Показывает модальное окно с сообщением и кнопкой OK."""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("350x140")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        # Центрируем диалог относительно окна
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 175
        y = self.winfo_y() + (self.winfo_height() // 2) - 70
        dialog.geometry(f"+{x}+{y}")

        ctk.CTkLabel(dialog, text=message, wraplength=320, justify="center").pack(pady=(20, 10))
        ctk.CTkButton(dialog, text="OK", width=90, command=dialog.destroy).pack(pady=(0, 15))

        # Быстрые клавиши
        dialog.bind("<Return>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        dialog.focus_set()

    def get_deal_id(self):
        """
        Возвращает ID сделки из поля ввода. Если пусто — спрашивает через модальное окно.
        При вводе через окно — заполняет поле ввода для удобства.
        """
        deal_id = self.deal_id_entry.get().strip()
        if not deal_id:
            deal_id = self.ask_deal_id()
            if deal_id:
                self.deal_id_entry.delete(0, "end")
                self.deal_id_entry.insert(0, deal_id)
            else:
                return None
        return deal_id

    def ask_deal_id(self):
        """Модальное окно для ввода ID сделки. Возвращает строку или None при отмене."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Введите ID сделки")
        dialog.geometry("320x140")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="ID сделки:", font=("Helvetica", 12)).pack(pady=(15, 5))
        entry = ctk.CTkEntry(dialog, width=220, font=("Helvetica", 12))
        entry.pack(pady=5)
        entry.focus()

        result = [None]

        def on_ok():
            value = entry.get().strip()
            result[0] = value if value else None
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="OK", command=on_ok, width=90).pack(side="left", padx=6)
        ctk.CTkButton(btn_frame, text="Отмена", command=on_cancel, width=90).pack(side="left", padx=6)

        # Ожидаем закрытия именно dialog (корректный способ)
        dialog.wait_window()
        return result[0]

    def show_error(self, msg):
        """Показывает диалог ошибки."""
        self.show_message("Ошибка", msg)

    def safe_call(self, func, deal_id, success_msg):
        """
        Безопасно вызывает функцию из модуля main.
        Показывает сообщение об успехе или ошибку.
        """
        try:
            func(deal_id)
            self.show_message("Готово", success_msg)
        except Exception as e:
            self.show_error(f"Ошибка: {e}")

    # === Обработчики кнопок ===
    def import_data(self):
        print("начали импорт...\n")
        if deal_id := self.get_deal_id():
            self.safe_call(self.main_module.import_data, deal_id, f"Данные для {deal_id} импортированы.")

    def calculate(self):
        if deal_id := self.get_deal_id():
            self.safe_call(self.main_module.calculate, deal_id, f"Расчёт для {deal_id} завершён.")

    def export_data(self):
        if deal_id := self.get_deal_id():
            self.safe_call(self.main_module.export_data, deal_id, "Результаты экспортированы.")

    def generate_3kp(self):
        if deal_id := self.get_deal_id():
            self.safe_call(self.main_module.generate_3kp, deal_id, f"3КП для {deal_id} создано.")

    def generate_tz(self):
        if deal_id := self.get_deal_id():
            self.safe_call(self.main_module.generate_tz, deal_id, f"ТЗ для {deal_id} создано.")

    def open_folder(self):
        """
        Открывает папку: [текущая]/Расчеты/{deal_id}
        Создаёт её при отсутствии.
        """
        deal_id = self.get_deal_id()
        if not deal_id:
            return

        current_dir = os.path.dirname(os.path.abspath(__file__))
        target_dir = os.path.join(current_dir, "Расчеты", deal_id)
        os.makedirs(target_dir, exist_ok=True)

        try:
            if sys.platform == "win32":
                os.startfile(target_dir)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", target_dir], check=False)
            else:
                subprocess.run(["xdg-open", target_dir], check=False)
        except Exception as e:
            self.show_error(f"Не удалось открыть папку:\n{e}")


if __name__ == "__main__":
    app = DealApp()
    app.mainloop()
