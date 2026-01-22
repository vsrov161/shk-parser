import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
import pandas as pd
from pathlib import Path


class JSONConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON Converter & Splitter")
        self.root.geometry("700x500")

        # Создаем вкладки
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Вкладка 1: Конвертация в JSON
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text='Конвертация в JSON')
        self.setup_conversion_tab()

        # Вкладка 2: Разделение JSON
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text='Разделение JSON')
        self.setup_split_tab()

        # Переменные
        self.input_file_path = None
        self.json_file_path = None

    def setup_conversion_tab(self):
        # Фрейм для выбора файла
        file_frame = ttk.LabelFrame(self.tab1, text="Выбор файла", padding=10)
        file_frame.pack(fill='x', padx=10, pady=(10, 5))

        ttk.Button(file_frame, text="Выбрать файл (XLS/XLSX/CSV)",
                   command=self.select_input_file).pack(pady=5)

        self.file_label = ttk.Label(file_frame, text="Файл не выбран",
                                    foreground="gray")
        self.file_label.pack(pady=5)

        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(self.tab1, text="Настройки", padding=10)
        settings_frame.pack(fill='x', padx=10, pady=5)

        # Параметры для JSON
        ttk.Label(settings_frame, text="employee_id:").grid(row=0, column=0, sticky='w', pady=2)
        self.employee_id_var = tk.StringVar(value="541143")
        ttk.Entry(settings_frame, textvariable=self.employee_id_var, width=20).grid(row=0, column=1, sticky='w', pady=2,
                                                                                    padx=(5, 10))

        ttk.Label(settings_frame, text="Дата и время (dt):").grid(row=1, column=0, sticky='w', pady=2)
        self.dt_var = tk.StringVar(value="2025-07-18T11:00:05.000000000+03:00")
        ttk.Entry(settings_frame, textvariable=self.dt_var, width=40).grid(row=1, column=1, sticky='w', pady=2,
                                                                           padx=(5, 10))

        ttk.Label(settings_frame, text="barcode:").grid(row=2, column=0, sticky='w', pady=2)
        self.barcode_var = tk.StringVar(value=" ")
        ttk.Entry(settings_frame, textvariable=self.barcode_var, width=20).grid(row=2, column=1, sticky='w', pady=2,
                                                                                padx=(5, 10))

        ttk.Label(settings_frame, text="task_type:").grid(row=3, column=0, sticky='w', pady=2)
        self.task_type_var = tk.StringVar(value="IKZ")
        ttk.Entry(settings_frame, textvariable=self.task_type_var, width=20).grid(row=3, column=1, sticky='w', pady=2,
                                                                                  padx=(5, 10))

        # Кнопка конвертации
        ttk.Button(self.tab1, text="Спарсить в JSON",
                   command=self.convert_to_json,
                   style="Accent.TButton").pack(pady=20)

        # Статус
        self.status_label = ttk.Label(self.tab1, text="")
        self.status_label.pack(pady=5)

    def setup_split_tab(self):
        # Фрейм для выбора JSON файла
        json_frame = ttk.LabelFrame(self.tab2, text="Выбор JSON файла", padding=10)
        json_frame.pack(fill='x', padx=10, pady=(10, 5))

        ttk.Button(json_frame, text="Выбрать JSON файл",
                   command=self.select_json_file).pack(pady=5)

        self.json_file_label = ttk.Label(json_frame, text="Файл не выбран",
                                         foreground="gray")
        self.json_file_label.pack(pady=5)

        # Фрейм для настроек разделения
        split_frame = ttk.LabelFrame(self.tab2, text="Настройки разделения", padding=10)
        split_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(split_frame, text="Количество элементов в каждом файле:").grid(row=0, column=0, sticky='w', pady=10)
        self.chunk_size_var = tk.StringVar(value="1000")
        ttk.Entry(split_frame, textvariable=self.chunk_size_var, width=15).grid(row=0, column=1, sticky='w', pady=10,
                                                                                padx=(10, 0))

        # Кнопка разделения
        ttk.Button(self.tab2, text="Разделить JSON файл",
                   command=self.split_json_file,
                   style="Accent.TButton").pack(pady=20)

        # Статус
        self.split_status_label = ttk.Label(self.tab2, text="")
        self.split_status_label.pack(pady=5)

    def select_input_file(self):
        filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]

        filename = filedialog.askopenfilename(
            title="Выберите файл с данными",
            filetypes=filetypes
        )

        if filename:
            self.input_file_path = filename
            self.file_label.config(
                text=f"Выбран: {os.path.basename(filename)}",
                foreground="black"
            )
            self.status_label.config(text="")

    def select_json_file(self):
        filetypes = [
            ("JSON files", "*.json"),
            ("All files", "*.*")
        ]

        filename = filedialog.askopenfilename(
            title="Выберите JSON файл",
            filetypes=filetypes
        )

        if filename:
            self.json_file_path = filename
            self.json_file_label.config(
                text=f"Выбран: {os.path.basename(filename)}",
                foreground="black"
            )
            self.split_status_label.config(text="")

    def read_file(self, file_path):
        """Чтение файла разных форматов"""
        try:
            ext = os.path.splitext(file_path)[1].lower()

            if ext in ['.xlsx', '.xls']:
                # Читаем Excel, предполагаем что данные в первом столбце
                df = pd.read_excel(file_path, header=None)
                # Берем первый столбец и фильтруем пустые значения
                values = df[0].dropna().astype(str).tolist()
                # Преобразуем в int если возможно, иначе оставляем как строку
                result = []
                for val in values:
                    try:
                        # Пробуем преобразовать в int
                        result.append(int(float(val)))
                    except:
                        # Если не число, оставляем как строку
                        result.append(val)
                return result

            elif ext == '.csv':
                # Читаем CSV, предполагаем что данные в первом столбце
                df = pd.read_csv(file_path, header=None, encoding='utf-8')
                values = df[0].dropna().astype(str).tolist()
                result = []
                for val in values:
                    try:
                        result.append(int(float(val)))
                    except:
                        result.append(val)
                return result
            else:
                messagebox.showerror("Ошибка", f"Неподдерживаемый формат файла: {ext}")
                return None

        except Exception as e:
            messagebox.showerror("Ошибка чтения файла", str(e))
            return None

    def convert_to_json(self):
        if not self.input_file_path:
            messagebox.showwarning("Внимание", "Сначала выберите файл!")
            return

        # Запрашиваем место сохранения
        save_path = filedialog.asksaveasfilename(
            title="Сохранить JSON файл",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not save_path:
            return

        try:
            # Читаем данные из файла
            goods_ids = self.read_file(self.input_file_path)

            if not goods_ids:
                return

            # Получаем параметры
            try:
                employee_id = int(self.employee_id_var.get())
            except:
                employee_id = self.employee_id_var.get()

            dt = self.dt_var.get()
            barcode = self.barcode_var.get()
            task_type = self.task_type_var.get()

            # Создаем массив объектов
            result = []
            for goods_id in goods_ids:
                result.append({
                    "employee_id": employee_id,
                    "goods_id": goods_id,
                    "dt": dt,
                    "barcode": barcode,
                    "task_type": task_type
                })

            # Сохраняем в JSON
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            self.status_label.config(
                text=f"Успешно! Создан файл: {os.path.basename(save_path)}\n"
                     f"Количество записей: {len(result)}",
                foreground="green"
            )

            messagebox.showinfo("Успех",
                                f"Файл успешно создан!\n"
                                f"Всего записей: {len(result)}\n"
                                f"Путь: {save_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.status_label.config(text=f"Ошибка: {str(e)}", foreground="red")

    def split_json_file(self):
        if not self.json_file_path:
            messagebox.showwarning("Внимание", "Сначала выберите JSON файл!")
            return

        try:
            # Получаем размер чанка
            chunk_size = int(self.chunk_size_var.get())
            if chunk_size <= 0:
                messagebox.showwarning("Внимание", "Размер чанка должен быть положительным числом!")
                return

        except ValueError:
            messagebox.showwarning("Внимание", "Введите корректное число для размера чанка!")
            return

        try:
            # Читаем JSON файл
            with open(self.json_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if not isinstance(data, list):
                messagebox.showerror("Ошибка", "JSON файл должен содержать массив!")
                return

            # Создаем папку для результатов
            base_name = os.path.splitext(os.path.basename(self.json_file_path))[0]
            output_dir = os.path.join(os.path.dirname(self.json_file_path), f"{base_name}_split")

            # Если папка уже существует, добавляем номер
            counter = 1
            original_dir = output_dir
            while os.path.exists(output_dir):
                output_dir = f"{original_dir}_{counter}"
                counter += 1

            os.makedirs(output_dir)

            # Разбиваем на части
            total_items = len(data)
            num_chunks = (total_items + chunk_size - 1) // chunk_size  # Округление вверх

            created_files = []
            for i in range(num_chunks):
                start_idx = i * chunk_size
                end_idx = min(start_idx + chunk_size, total_items)
                chunk = data[start_idx:end_idx]

                # Сохраняем чанк
                chunk_filename = f"{base_name}_part_{i + 1:03d}.json"
                chunk_path = os.path.join(output_dir, chunk_filename)

                with open(chunk_path, 'w', encoding='utf-8') as f:
                    json.dump(chunk, f, ensure_ascii=False, indent=2)

                created_files.append(chunk_path)

            # Показываем результат
            self.split_status_label.config(
                text=f"Успешно! Создано {num_chunks} файлов\n"
                     f"Папка: {os.path.basename(output_dir)}",
                foreground="green"
            )

            messagebox.showinfo("Успех",
                                f"Файл успешно разделен!\n"
                                f"Всего частей: {num_chunks}\n"
                                f"Элементов в каждой части: ~{chunk_size}\n"
                                f"Папка с результатами: {output_dir}")

        except json.JSONDecodeError:
            messagebox.showerror("Ошибка", "Некорректный JSON формат файла!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            self.split_status_label.config(text=f"Ошибка: {str(e)}", foreground="red")


def main():
    root = tk.Tk()

    # Настройка стиля для кнопки акцента
    style = ttk.Style()
    style.configure("Accent.TButton", foreground="white", background="#0078d7")

    app = JSONConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()