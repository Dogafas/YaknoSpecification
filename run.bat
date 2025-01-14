@echo off
echo Starting Spec Generator...

:: Проверка наличия виртуального окружения
if not exist venv (
  echo Virtual environment "venv" not found. Please create it using: python -m venv venv
  goto :eof
)

:: Активация виртуального окружения
call venv\Scripts\activate

:: Проверка на наличие файла requirements.txt
if not exist requirements.txt (
  echo requirements.txt not found.  Please make sure it's in the same directory.
  goto :deactivate
)

:: Установка пакетов (только если файл requirements.txt существует)
pip install -r requirements.txt

:: Запуск программы
python start.py

:deactivate
echo Deactivating virtual environment...
call venv\Scripts\deactivate
echo Finished.
pause