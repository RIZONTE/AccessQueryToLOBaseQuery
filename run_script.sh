#!/bin/bash

# Пути и настройки
PYTHON_SCRIPT="get_csv_excel.py"

echo "Активация виртуального окружения..."
source "$VENV_DIR/bin/activate"

# Запуск Python-скрипта
echo "Запуск $PYTHON_SCRIPT..."
python3 "$PYTHON_SCRIPT"

# Конец
echo "Выполнение завершено."