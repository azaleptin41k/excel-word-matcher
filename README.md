# Excel-Word Cross-Matcher & Formatter

[![CI & Security Pipeline](https://github.com/azaleptin41k/excel-word-matcher/actions/workflows/ci-security.yml/badge.svg?branch=master)](https://github.com/azaleptin41k/excel-word-matcher/actions/workflows/ci-security.yml)

Утилита с графическим интерфейсом для автоматического парсинга Word-документов,
сверки данных с Excel-таблицами и условного цветового форматирования строк.
Создана для автоматизации работы отдела закупок.

---

## Стек технологий

| Слой | Инструмент |
|---|---|
| Язык | Python 3.10+ |
| GUI | `tkinter` (стандартная библиотека) |
| Работа с Excel | `openpyxl` |
| Работа с Word | `python-docx` |
| Сборка в EXE | `PyInstaller` |
| Контейнеризация | Docker (multi-stage build) |
| CI/CD | GitHub Actions |
| SAST | Bandit |
| Secret Scanning | Gitleaks |
| Dependency Audit | pip-audit |
| Container Scanning | Trivy |
| SBOM | Syft (SPDX) |
| Linting | Ruff |
| Тестирование | pytest + coverage |

---

## Security Pipeline

Проект включает полноценный DevSecOps pipeline, запускаемый автоматически при каждом push и PR:

```
push / PR
  │
  ├─ Lint (Ruff)                  — проверка стиля и потенциальных ошибок
  ├─ Tests (pytest)               — юнит-тесты с покрытием
  ├─ SAST (Bandit)                — статический анализ безопасности Python-кода
  ├─ Secret Scan (Gitleaks)       — обнаружение случайно закоммиченных секретов
  ├─ Dependency Audit (pip-audit) — проверка зависимостей на известные CVE
  ├─ Docker Build & Scan (Trivy)  — сборка образа + сканирование на уязвимости
  ├─ SBOM (Syft)                  — генерация Software Bill of Materials
  │
  └─ Security Gate                — финальная проверка: все критические шаги пройдены
```

Подробнее о политике безопасности — в [SECURITY.md](SECURITY.md).

---

## Возможности

1. **Извлечение данных** — считывает указанные столбцы Excel-файла начиная с заданной строки и сохраняет их во временный текстовый файл.
2. **Поиск по кодам** — сверяет коды из пользовательского `.txt`-файла с извлечёнными данными (поддерживаются иерархические коды ОКПД2 с учётом вложенности).
3. **Цветовое выделение** — подсвечивает строки Excel выбранным цветом, если значение в ключевом столбце совпадает с найденными кодами.
4. **Заполнение категорий** — парсит таблицу из `categories.docx` и проставляет названия категорий в нужный столбец Excel по найденным кодам.

---

## Структура проекта

```
excel-word-matcher/
├── .github/
│   └── workflows/
│       └── ci-security.yml     # CI/CD pipeline (lint, SAST, secrets, Trivy, SBOM)
├── tests/
│   ├── __init__.py
│   └── test_core.py            # Юнит-тесты ядра (extract, match, highlight)
├── beaver2.py                  # Основной скрипт (GUI + вся логика)
├── beaver2.spec                # Конфигурация PyInstaller
├── categories.docx             # Справочник категорий ОКПД2 (Word-таблица, 2 столбца)
├── dummy_data.xlsx             # Тестовый Excel-файл с 65 строками фиктивных закупок
├── Dockerfile                  # Multi-stage сборка (builder + non-root runtime)
├── .dockerignore
├── .gitignore
├── .gitleaks.toml              # Конфигурация Gitleaks
├── .pre-commit-config.yaml     # Pre-commit hooks (Ruff, Bandit, Gitleaks, и др.)
├── pyproject.toml              # Настройки Ruff, pytest, Bandit
├── requirements.txt            # Runtime-зависимости
├── requirements-dev.txt        # Dev-зависимости (pytest, ruff, bandit, и др.)
└── SECURITY.md                 # Политика безопасности
```

### Входные файлы

| Файл | Назначение |
|---|---|
| Исходный Excel (любое имя) | Таблица закупок; выбирается через GUI |
| `.txt`-файл с кодами | Коды ОКПД2 для поиска (по одному на строку) |
| `categories.docx` | Справочник `[код → название категории]`; таблица Word, 2 столбца |

### Выходной файл

| Файл | Назначение |
|---|---|
| `output.xlsx` | Копия исходного Excel с цветовой разметкой и (опционально) заполненными категориями |

---

## Быстрый старт (запуск из исходников)

```bash
# 1. Установить зависимости
pip install -r requirements.txt

# 2. Запустить приложение
python beaver2.py
```

### Порядок работы

1. **Выбрать исходный Excel** — файл будет скопирован как `wrong_data.xlsx`.
2. Указать **столбцы** для извлечения (по умолчанию `E,C`) и **начальную строку** (по умолчанию `9`).
3. **Выбрать файл с кодами** (`.txt`) — каждый код на новой строке.
4. Выбрать **цвет** для подсветки строк.
5. Нажать **«Запустить обработку»** → будет создан `output.xlsx`.
6. (Опционально) Выбрать файл для ОКПД2, указать столбцы поиска/вывода и нажать **«Заполнить ОКПД2»**.

> **Подсказка:** В папке `categories.docx` уже есть готовый справочник из 25 категорий.
> `dummy_data.xlsx` содержит 65 строк тестовых закупок — используйте его для проверки.

---

## Разработка

```bash
# Установить все зависимости (runtime + dev)
pip install -r requirements.txt -r requirements-dev.txt

# Настроить pre-commit hooks
pre-commit install

# Запустить тесты
pytest tests/ -v --cov=beaver2

# Запустить линтер
ruff check .

# Запустить SAST вручную
bandit -r . --severity-level medium
```

---

## Сборка исполняемого файла (.exe)

```bash
pyinstaller beaver2.spec
```

Готовый `.exe` появится в папке `dist/beaver2/`.

> Файл `categories.docx` включён в сборку автоматически (прописан в `beaver2.spec`).

---

## Docker (Multi-stage сборка)

Dockerfile использует **multi-stage build** для минимального размера образа и **non-root user** для безопасного запуска.

Поскольку сборка идёт на базе Linux-образа, на выходе генерируется исполняемый бинарный файл Linux (ELF). *Для сборки Windows .exe через Docker потребовался бы образ на базе Wine.*

```bash
# 1. Собрать образ
docker build -t excel-word-matcher .

# 2. Запустить сборку с пробросом папки dist на хост-машину
# Для Windows (PowerShell):
docker run --rm -v ${PWD}/dist:/app/dist excel-word-matcher

# Для Linux/Mac/Git Bash:
docker run --rm -v "$(pwd)/dist:/app/dist" excel-word-matcher
```

После завершения команды в вашей локальной папке `dist/` появится собранный бинарник.

---

## Настройка исключения кодов

В `beaver2.py` есть константа `EXCLUDED_OKPD2_CODE` — код ОКПД2, строки с которым
**не** будут подсвечиваться даже при совпадении:

```python
EXCLUDED_OKPD2_CODE = "XX.XX.XX.XXX"  # замените на нужный код
```

