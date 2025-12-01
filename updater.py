# updater.py

import sys
import time
import shutil
import subprocess

from email.utils import parsedate_to_datetime
from urllib.parse import urljoin, quote
import requests
import base64
import os
from pathlib import Path

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

    def upload(self, token: str, owner: str, repo: str, branch: str = "main"):
        # --- Проверка токена и доступа к репозиторию ---
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "X-GitHub-Api-Version": "2022-11-28"
        }
        repo_url = f"https://api.github.com/repos/{owner}/{repo}"

        print("🔍 Проверка доступа к репозиторию...")
        resp = requests.get(repo_url, headers=headers)
        if resp.status_code == 404:
            print(f"❌ Репозиторий не найден: {owner}/{repo}")
            print("   Убедитесь, что:")
            print("   - Имя владельца и репозитория написаны правильно")
            print("   - Репозиторий существует")
            print("   - Если репозиторий приватный — токен имеет доступ")
            return
        elif resp.status_code == 401:
            print("❌ Недействительный или отсутствующий токен.")
            print(
                "   Убедитесь, что GITHUB_TOKEN задан и имеет scope 'repo' (или 'public_repo' для публичных репозиториев).")
            return
        elif resp.status_code != 200:
            print(f"❌ Ошибка доступа к репозиторию ({resp.status_code}): {resp.json().get('message', 'Unknown error')}")
            return

        repo_info = resp.json()
        print(f"✅ Доступ к репозиторию подтверждён: {repo_info['full_name']}")
        # if not repo_info.get("permissions", {}).get("push", False):
        #     print("⚠️  Внимание: у токена нет прав на запись (push) в репозиторий!")
        #     print("   Обновление файлов не удастся.")
        #     return
        print(repo_info)

        # --- Проверка существования ветки ---
        branches_url = f"https://api.github.com/repos/{owner}/{repo}/branches/{branch}"
        branch_resp = requests.get(branches_url, headers=headers)
        if branch_resp.status_code == 404:
            print(f"❌ Ветка '{branch}' не существует в репозитории.")
            print("   Убедитесь, что имя ветки указано верно (по умолчанию: 'main' или 'master').")
            return
        elif branch_resp.status_code != 200:
            print(f"⚠️  Не удалось проверить ветку '{branch}' ({branch_resp.status_code})")
        else:
            print(f"✅ Ветка '{branch}' существует.")

        print("➡️  Приступаю к загрузке файлов...\n")

        api_base = f"https://api.github.com/repos/{owner}/{repo}/contents/"  # ← исправлено!

        for file in self.files_to_update:
            local_path = Path(self.local_dir) / file
            if not local_path.exists():
                print(f"⚠ {file}: не найден локально, пропускаем")
                continue

            with open(local_path, "rb") as f:
                content = f.read()
            encoded_content = base64.b64encode(content).decode("utf-8")

            # Нормализуем и кодируем путь
            remote_path = quote(str(Path(file).as_posix()), safe="/")
            url = api_base + remote_path
            params = {"ref": branch}

            # Получаем текущий SHA, если файл существует
            resp = requests.get(url, headers=headers, params=params)
            data = {
                "message": f"Update {file} via updater.py",
                "content": encoded_content,
                "branch": branch
            }

            if resp.status_code == 200:
                data["sha"] = resp.json()["sha"]
            elif resp.status_code != 404:
                print(f"❌ {file}: ошибка при проверке существования ({resp.status_code}) — {resp.text}")
                continue

            # Загружаем/обновляем
            upload_resp = requests.put(url, headers=headers, json=data)
            if upload_resp.status_code in (200, 201):
                print(f"✅ {file} успешно загружен в {owner}/{repo}")
            else:
                print(f"❌ {file}: ошибка загрузки ({upload_resp.status_code}) — {upload_resp.json()}")

def update():
    import os
    from config import UPDATE_BASE_URL, FILES_TO_UPDATE, LOCAL_APP_DIR
    updater = Updater(UPDATE_BASE_URL, FILES_TO_UPDATE, LOCAL_APP_DIR)
    updater.auto_update_check()

def upload():
    import os
    from config import UPDATE_BASE_URL, FILES_TO_UPDATE, LOCAL_APP_DIR, GITHUB_TOKEN
    updater = Updater(UPDATE_BASE_URL, FILES_TO_UPDATE, LOCAL_APP_DIR)
    updater.upload(
        token=GITHUB_TOKEN,
        owner="Latortsev",
        repo="LOA_LSO",
        branch="main"
    )


if __name__ == "__main__":
    upload()



