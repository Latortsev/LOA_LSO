@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo Проверка наличия Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo Python не найден. Устанавливаю последнюю версию Python...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe' -OutFile 'python-installer.exe'"
    echo Запуск установки Python...
    python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    del python-installer.exe
    echo Установка завершена. Перезапуск...
    python --version
) else (
    echo Python уже установлен.
)

echo Установка зависимостей из requirements.txt...
pip install -r requirements.txt

echo Создание ярлыка на рабочем столе...
powershell -ExecutionPolicy Bypass -Command ^
    "$WshShell = New-Object -comObject WScript.Shell; " ^
    "$Shortcut = $WshShell.CreateShortcut([System.Environment]::GetFolderPath('Desktop') + '\\Расчет_ЛШО.lnk'); " ^
    "$Shortcut.TargetPath = 'pythonw.exe'; " ^
    "$Shortcut.Arguments = '%~dp0gui.pyw'; " ^
    "$Shortcut.IconLocation = '%SystemRoot%\\system32\\imageres.dll,70'; " ^
    "$Shortcut.Description = 'Расчет ЛШО'; " ^
    "$Shortcut.Save();"

if exist "%%USERPROFILE%%\Desktop\Расчет_ЛШО.lnk" (
    echo Ярлык создан на рабочем столе.
) else (
    echo Ошибка: ярлык не был создан.
)

pause
