import tkinter as tk
from tkinter import filedialog, colorchooser, messagebox
import os
import shutil
from openpyxl import load_workbook
import re
from docx import Document
from openpyxl.utils import column_index_from_string
from openpyxl.styles import PatternFill

# Код, строки с которым исключаются из цветового выделения.
# Замените на реальный код из вашей классификации.
EXCLUDED_OKPD2_CODE = "XX.XX.XX.XXX"

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик Excel")
        self.root.geometry("600x600")

        # Параметры getdata.py
        self.frame_getdata = tk.LabelFrame(root, text="1. Настройка извлечения данных")
        self.frame_getdata.pack(padx=10, pady=10, fill="x")

        self.columns_var = tk.StringVar(value="E,C")
        tk.Label(self.frame_getdata, text="Столбцы (через запятую):").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(self.frame_getdata, textvariable=self.columns_var).grid(row=0, column=1, padx=5, pady=5)

        self.start_row_var = tk.IntVar(value=9)
        tk.Label(self.frame_getdata, text="Начальная строка:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(self.frame_getdata, textvariable=self.start_row_var).grid(row=1, column=1, padx=5, pady=5)

        self.source_file = tk.StringVar()
        tk.Button(self.frame_getdata, text="Выбрать исходный Excel",
                 command=self.select_source).grid(row=2, column=0, columnspan=2, pady=5)

        # Параметры finder1.02.py
        self.frame_finder = tk.LabelFrame(root, text="2. Настройка поиска")
        self.frame_finder.pack(padx=10, pady=10, fill="x")

        self.numbers_file = tk.StringVar()
        tk.Button(self.frame_finder, text="Выбрать файл номеров",
                 command=self.select_numbers).pack(padx=5, pady=5)

        # Параметры finderexcele
        self.frame_highlight = tk.LabelFrame(root, text="3. Настройка выделения")
        self.frame_highlight.pack(padx=10, pady=10, fill="x")

        self.color_var = tk.StringVar(value="#44944a")
        tk.Button(self.frame_highlight, text="Выбрать цвет выделения",
                 command=self.choose_color).pack(padx=5, pady=5)

        # Дополнительные настройки
        self.use_ogr_var = tk.BooleanVar()
        tk.Checkbutton(root, text="Использовать ogrfinderexcele",
                     variable=self.use_ogr_var).pack(padx=10, pady=5)

        # Кнопки управления
        self.btn_frame = tk.Frame(root)
        self.btn_frame.pack(padx=10, pady=20)

        tk.Button(self.btn_frame, text="Запустить обработку",
                 command=self.process).pack(side=tk.LEFT, padx=5)
        tk.Button(self.btn_frame, text="Заполнить ОКПД2",
                 command=self.run_okpd).pack(side=tk.RIGHT, padx=5)

        # Для names_numbers.py
        self.okpd_frame = tk.LabelFrame(root, text="Дополнительные настройки ОКПД2")
        self.okpd_frame.pack(padx=10, pady=10, fill="x")
        self.okpd_file = tk.StringVar()
        tk.Button(self.okpd_frame, text="Выбрать файл для ОКПД2",
                 command=self.select_okpd_file).grid(row=2, column=0, columnspan=2, pady=5)

        self.search_col_var = tk.StringVar(value="C")
        self.output_col_var = tk.StringVar(value="D")

        tk.Label(self.okpd_frame, text="Столбец поиска:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(self.okpd_frame, textvariable=self.search_col_var).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.okpd_frame, text="Столбец вывода:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(self.okpd_frame, textvariable=self.output_col_var).grid(row=1, column=1, padx=5, pady=5)

    def select_source(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.source_file.set(file_path)
            # Автоматическое переименование
            try:
                shutil.copy(file_path, "wrong_data.xlsx")
                messagebox.showinfo("Успех", "Файл скопирован как wrong_data.xlsx")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def select_numbers(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            self.numbers_file.set(file_path)

    def choose_color(self):
        color = colorchooser.askcolor()[1]
        if color:
            self.color_var.set(color)

    def select_okpd_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.okpd_file.set(file_path)

    def process_getdata(self):
        # Логика getdata.py
        COLUMNS_TO_COPY = self.columns_var.get().split(',')
        START_ROW = self.start_row_var.get()
        SOURCE_FILE = "wrong_data.xlsx"
        OUTPUT_FILE = "vse.txt"

        try:
            src_wb = load_workbook(SOURCE_FILE)
            src_ws = src_wb.active
            column_indices = [column_index_from_string(col.strip()) - 1 for col in COLUMNS_TO_COPY]

            with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
                for row in src_ws.iter_rows(min_row=START_ROW, values_only=True):
                    current_values = [row[idx] for idx in column_indices]
                    if all(v is None for v in current_values):
                        continue
                    formatted_values = [str(v) if v is not None else '' for v in current_values]
                    f.write('\t'.join(formatted_values) + '\n')
            return True
        except Exception as e:
            messagebox.showerror("Ошибка getdata", str(e))
            return False

    def process_finder(self):
        # Логика finder1.02.py
        numbers_file = self.numbers_file.get()
        if not os.path.exists(numbers_file):
            messagebox.showerror("Ошибка", f"Файл {numbers_file} не найден")
            return False

        try:
            def read_numbers_from_file(filename):
                with open(filename, 'r', encoding='utf-8') as f:
                    return [line.strip() for line in f if line.strip()]

            def read_elements_from_file(filename):
                with open(filename, 'r', encoding='utf-8') as f:
                    return [line.strip() for line in f if line.strip()]

            numbers = read_numbers_from_file(numbers_file)
            elements = read_elements_from_file('vse.txt')
            output_file = 'found_elements.txt'

            with open(output_file, 'w', encoding='utf-8') as file:
                for number in numbers:
                    pattern = rf'^{re.escape(number)}(\.|$)'
                    parts = number.split('.')
                    last_part = parts[-1]
                    
                    # --- НАЧАЛО ИЗМЕНЕНИЯ ---
                    # Новое условие: проверяем, если последняя часть - одна цифра ИЛИ если она заканчивается на '0' (и не является просто '0')
                    if len(last_part) == 1 or (last_part.endswith('0') and len(last_part) > 1):
                        
                        # Если номер заканчивается на 0 (например, 12.34.56.10), 
                        # мы должны использовать '12.34.56.1' как основу для поиска.
                        if last_part.endswith('0'):
                            base_number_prefix = '.'.join(parts[:-1] + [last_part[:-1]])
                        else:
                            # Если последняя часть - одна цифра (например, 12.34.56.1),
                            # используем номер как есть.
                            base_number_prefix = number

                        search_prefix = f"{base_number_prefix}"
                        for element in elements:
                            if not element:
                                continue
                            
                            # Добавлена обработка ошибок на случай пустых строк в element
                            try:
                                left_part = element.split()[0]
                                second_part = element.split()[-1]
                            except IndexError:
                                continue

                            if any(second_part.startswith(f"{search_prefix}{i}") for i in range(10)):
                                file.write(f"{left_part}\n")
                    
                    # --- КОНЕЦ ИЗМЕНЕНИЯ ---

                    for element in elements:
                        if not element:
                            continue
                        
                        # Добавлена обработка ошибок на случай пустых строк в element
                        try:
                            left_part = element.split()[0]
                            second_part = element.split()[-1]
                        except IndexError:
                            continue

                        if re.match(pattern, second_part):
                            file.write(f"{left_part}\n")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка finder", str(e))
            return False

    def process_highlight(self):
        # Логика finderexcele/ogrfinderexcele
        excel_file = "wrong_data.xlsx"
        output_file = "output.xlsx"
        column_to_search = self.columns_var.get().split(',')[0]
        print(column_to_search)
        color = self.color_var.get()

        if self.use_ogr_var.get():
            # Логика ogrfinderexcele_2.py
            def highlight_matching_rows(workbook, elements):
                sheet = workbook.active
                green_fill = PatternFill(start_color=color[1:],
                                        end_color=color[1:],
                                        fill_type="solid")
                elements_set = set(elements)
                col_idx = column_index_from_string(column_to_search) - 1

                for row in sheet.iter_rows():
                    if len(row) <= col_idx:
                        continue
                    cell = row[col_idx]
                    cell_value = str(cell.value).strip() if cell.value else ""
                    if cell_value in elements_set:
                        has_forbidden = any(
                            str(c.value).strip() == EXCLUDED_OKPD2_CODE
                            for c in row
                        )
                        if not has_forbidden:
                            for c in row:
                                c.fill = green_fill
        else:
            # Логика finderexcele_2.py
            def highlight_matching_rows(workbook, elements):
                sheet = workbook.active
                green_fill = PatternFill(start_color=color[1:],
                                        end_color=color[1:],
                                        fill_type="solid")
                elements_set = set(elements)
                col_idx = column_index_from_string(column_to_search) - 1

                for row in sheet.iter_rows():
                    if len(row) <= col_idx:
                        continue
                    cell = row[col_idx]
                    cell_value = str(cell.value).strip() if cell.value else ""
                    if cell_value in elements_set:
                        for c in row:
                            c.fill = green_fill

        try:
            elements = [line.strip() for line in open('found_elements.txt', 'r', encoding='utf-8')]
            wb = load_workbook(excel_file)
            highlight_matching_rows(wb, elements)
            wb.save(output_file)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка выделения", str(e))
            return False

    def cleanup_temp_files(self):
        temp_files = [
            "wrong_data.xlsx",
            "vse.txt",
            "found_elements.txt"
        ]
        for file in temp_files:
            try:
                if os.path.exists(file):
                    os.remove(file)
            except Exception as e:
                messagebox.showwarning("Предупреждение", f"Не удалось удалить {file}: {str(e)}")

    def process(self):
        try:
            # 1. Выполняем getdata
            if not self.process_getdata():
                return

            # 2. Выполняем finder
            if not self.process_finder():
                return

            # 3. Выполняем выделение
            if not self.process_highlight():
                return

            messagebox.showinfo("Успех", "Обработка завершена успешно!")
        except Exception as e:
            messagebox.showerror("Общая ошибка", str(e))
        finally:
            self.cleanup_temp_files()

    def run_okpd(self):
        # Проверяем, выбран ли файл
        if not self.okpd_file.get():
            messagebox.showerror("Ошибка ОКПД2", "Файл для ОКПД2 не выбран")
            return

        src_file = self.okpd_file.get()
        if os.path.basename(src_file) != "output.xlsx":
            try:
                shutil.copy(src_file, "output.xlsx")
            except Exception as e:
                messagebox.showerror("Ошибка ОКПД2", f"Не удалось скопировать файл: {str(e)}")
                return
            else:
                # Если файл уже называется output.xlsx, используем его напрямую
                pass

        # Логика names_numbers.py
        input_excel = "output.xlsx"
        input_word = "categories.docx"
        search_col = self.search_col_var.get()
        output_col = self.output_col_var.get()
        start_row = self.start_row_var.get()

        def get_merged_cell(ws, row, col):
            for merged in ws.merged_cells.ranges:
                if (merged.min_row <= row <= merged.max_row and
                    merged.min_col <= col <= merged.max_col):
                    return ws.cell(merged.min_row, merged.min_col)
            return None

        try:
            wb = load_workbook(input_excel)
            ws = wb.active
            doc = Document(input_word)
            word_map = {}

            for table in doc.tables:
                for row in table.rows:
                    if len(row.cells) >= 2:
                        key = row.cells[0].text.strip()
                        val = row.cells[1].text.strip()
                        word_map[key] = val

            search_col_idx = column_index_from_string(search_col) - 1
            output_col_idx = column_index_from_string(output_col) - 1

            for row in ws.iter_rows(min_row=start_row):
                if len(row) <= search_col_idx:
                    continue
                cell = row[search_col_idx]
                search_val = str(cell.value).strip()
                if not search_val:
                    continue
                match_val = word_map.get(search_val, "")

                target_row = cell.row
                target_col = output_col_idx + 1

                merged = get_merged_cell(ws, target_row, target_col)
                if merged:
                    merged.value = match_val
                else:
                    ws.cell(row=target_row, column=target_col, value=match_val)

            wb.save("output.xlsx")
            messagebox.showinfo("Успех", "ОКПД2 заполнены!")
        except Exception as e:
            messagebox.showerror("Ошибка ОКПД2", str(e))
        finally:
            # Не удаляем output.xlsx после этой операции
            temp_files = ["wrong_data.xlsx", "vse.txt", "found_elements.txt"]
            for file in temp_files:
                try:
                    if os.path.exists(file):
                        os.remove(file)
                except Exception as e:
                    messagebox.showwarning("Предупреждение", f"Не удалось удалить {file}: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()
