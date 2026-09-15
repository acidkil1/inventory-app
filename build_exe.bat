@echo off
echo Создание EXE файла из приложения Inventory App...
echo.

REM Проверка наличия PyInstaller
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Установка PyInstaller...
    python -m pip install pyinstaller
)

REM Создание EXE файла с иконкой
echo.
echo Создание EXE файла с иконкой inventory.ico...
pyinstaller --onefile --windowed --icon=inventory.ico --name="Inventory" --add-data "db_manager.py;." --add-data "inventory.db;." --add-data "inventory.ico;." inventory.py

echo.
echo Готово! EXE файл находится в папке dist\
echo.
echo Для запуска приложения скопируйте следующие файлы в одну папку с exe:
echo - inventory.db (база данных)
echo - db_manager.py (модуль базы данных)
echo.
pause