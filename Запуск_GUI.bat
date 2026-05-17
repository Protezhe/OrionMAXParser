@echo off
setlocal

:: Check if local virtual environment exists
if exist ".venv\Scripts\python.exe" (
    echo [OK] Найдено виртуальное окружение .venv
    ".venv\Scripts\python.exe" gui_app.py
) else if exist "venv\Scripts\python.exe" (
    echo [OK] Найдено виртуальное окружение venv
    "venv\Scripts\python.exe" gui_app.py
) else (
    echo [!] Виртуальное окружение не найдено. Используется системный python.
    python gui_app.py
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo [!] Ошибка при запуске. Проверьте, установлен ли Python и зависимости.
    pause
)
