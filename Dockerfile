# Используем официальный легковесный образ Python
FROM python:3.10-slim

# Устанавливаем системные зависимости, необходимые для Tkinter (gui) и PyInstaller (сборка)
RUN apt-get update && apt-get install -y \
    python3-tk \
    binutils \
    && rm -rf /var/lib/apt/lists/*

# Создаем рабочую директорию
WORKDIR /app

# Копируем файлы зависимостей
COPY requirements.txt .

# Устанавливаем Python-зависимости (включая pyinstaller из requirements.txt)
RUN pip install --no-cache-dir -r requirements.txt

# Копируем исходный код проекта, спецификацию PyInstaller и справочник
COPY beaver2.py beaver2.spec categories.docx ./

# Точка входа по умолчанию: запуск сборки PyInstaller
# Скомпилированный бинарник появится во внутренней папке /app/dist/
CMD ["pyinstaller", "--clean", "-y", "beaver2.spec"]