# updater.py
import os
import sys
import time
import shutil
import subprocess
from pathlib import Path
from urllib.parse import urljoin
from email.utils import parsedate_to_datetime

import requests


class Updater:
    def __init__(self, base_url: str, files_to_update: list[str], local_dir: str):
        self.base_url = base_url
        self.files_to_update = files_to_update
        self.local_dir = local_dir

    def archive_existing_files(self, target_folder: str, keep_last: int = 5):
        """Архивирует файлы в подпапку с порядковым номером. Хранит только последние N архивов."""
        target_path = Path(target_folder)
        target_path.mkdir(exist_ok=True)

        files = [f for f in target_path.iterdir() if f.is_file()]
        if not files:
            return

        existing_archives = [d for d in target_path.iterdir() if d.is_dir() and d.name.isdigit()]
        next_archive_num = max([int(d.name) for d in existing_archives], default=0) + 1
        archive_folder = target_path / str(next_archive_num)
        archive_folder.mkdir()

        for file in files:
            shutil.move(str(file), str(archive_folder))

        # ограничиваем количество архивов
        if len(existing_archives) + 1 > keep_last:
            to_delete = sorted(existing_archives, key=lambda d: int(d.name))[:-keep_last+1]
            for d in to_delete:
                shutil.rmtree(d, ignore_errors=True)

    def check_for_updates(self) -> bool:
        """Проверяет наличие обновлений на сервере (по размеру или дате)."""
        for file in self.files_to_update:
            remote_url = urljoin(self.base_url, file)
            local_path = os.path.join(self.local_dir, file)

            try:
                response = requests.get(remote_url, stream=True)
                if response.status_code != 200:
                    print(f"❌ {file}: сервер вернул {response.status_code}")
                    continue

                # если файла нет локально → обновление
                if not os.path.exists(local_path):
                    print(f"⚠ {file} отсутствует локально → обновление нужно")
                    return True

                # сравнение по размеру
                remote_size = int(response.headers.get("Content-Length", 0))
                local_size = os.path.getsize(local_path)
                if remote_size and remote_size != local_size:
                    print(f"⚠ {file}: размер отличается (локально {local_size}, сервер {remote_size})")
                    return True

                # сравнение по дате (если сервер отдал Last-Modified)
                remote_time_str = response.headers.get("Last-Modified")
                if remote_time_str:
                    remote_time = parsedate_to_datetime(remote_time_str).timestamp()
                    local_time = os.path.getmtime(local_path)
                    if remote_time > local_time:
                        print(f"⚠ {file}: сервер новее (локально {time.ctime(local_time)}, сервер {remote_time_str})")
                        return True

            except Exception as e:
                print(f"Ошибка при проверке {file}: {e}")

        return False

    def update_files(self) -> bool:
        """Скачивает и обновляет файлы."""
        for file in self.files_to_update:
            remote_url = urljoin(self.base_url, file)
            local_path = os.path.join(self.local_dir, file)
            try:
                response = requests.get(remote_url)
                if response.status_code == 200:
                    os.makedirs(os.path.dirname(local_path), exist_ok=True)
                    with open(local_path, "wb") as f:
                        f.write(response.content)
                    print(f"✅ {file} обновлён ({len(response.content)} байт)")
                else:
                    print(f"❌ {file}: ошибка {response.status_code}")
                    return False
            except Exception as e:
                print(f"Ошибка скачивания {file}: {e}")
                return False
        return True

    def restart_app(self):
        """Перезапускает приложение."""
        print("Перезапуск приложения...")
        subprocess.Popen([sys.executable] + sys.argv)
        sys.exit()

    def auto_update_check(self):
        """Полный цикл проверки и обновления."""
        print("=== НАЧАЛО ПРОВЕРКИ ОБНОВЛЕНИЙ ===")
        for file in self.files_to_update:
            local_path = os.path.join(self.local_dir, file)
            if os.path.exists(local_path):
                print(f"✔ {file} существует ({os.path.getsize(local_path)} байт)")
            else:
                print(f"⚠ {file} отсутствует")

        if self.check_for_updates():
            print("🔄 Обнаружено обновление, начинаем загрузку...")
            if self.update_files():
                print("Обновление завершено, перезапуск...")
                self.restart_app()
            else:
                print("❌ Ошибка обновления.")
        else:
            print("Обновлений не обнаружено.")
        print("=== КОНЕЦ ПРОВЕРКИ ОБНОВЛЕНИЙ ===")

if __name__ == "__main__":
    # Example usage

    import os

    UPDATE_BASE_URL = "https://bitrix24public.com/labkabinet.bitrix24.ru/docs/pub/a74e057419b211005403b334135e4de9/default/"
    FILES_TO_UPDATE = [
        "main.py",
        "gui.pyw",
        "install.bat",
        "Шаблоны/Расчет_шаблон_V1.xlsx",
        "README.docx",
        "requirements.txt"
    ]
    LOCAL_APP_DIR = os.path.dirname(os.path.abspath(__file__))

    updater = Updater(UPDATE_BASE_URL, FILES_TO_UPDATE, LOCAL_APP_DIR)
    updater.auto_update_check()

