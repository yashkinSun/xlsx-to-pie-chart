import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import datetime
from data_loader import DataLoader
from data_analyzer import DataAnalyzer
from data_visualizer import DataVisualizer
import threading

class NonconformanceAnalyzerApp:
    """
    Приложение для анализа несоответствий на производстве.
    """
    
    def __init__(self, root):
        """
        Инициализация приложения.
        
        Args:
            root (tk.Tk): Корневой виджет Tkinter
        """
        self.root = root
        self.root.title("Анализатор несоответствий")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # Инициализация компонентов
        self.loader = DataLoader()
        self.analyzer = DataAnalyzer()
        self.visualizer = DataVisualizer()
        
        # Текущие данные
        self.current_data = None
        self.current_results = None
        self.current_figure = None
        self.comparison_figure = None
        
        # Создание интерфейса
        self._create_ui()
        
        # Журнал событий
        self.log("Приложение запущено")
    
    def _create_ui(self):
        """
        Создание пользовательского интерфейса.
        """
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Верхняя панель с кнопками
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Кнопка загрузки файла
        self.upload_button = ttk.Button(
            top_frame, 
            text="Загрузить Excel файл", 
            command=self._upload_file
        )
        self.upload_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Флажок для сравнения с предыдущим месяцем
        self.compare_var = tk.BooleanVar(value=False)
        self.compare_checkbox = ttk.Checkbutton(
            top_frame,
            text="Сравнить с предыдущим месяцем",
            variable=self.compare_var,
            command=self._toggle_comparison
        )
        self.compare_checkbox.pack(side=tk.LEFT, padx=(0, 10))
        
        # Кнопка скачивания отчета
        self.download_button = ttk.Button(
            top_frame,
            text="Скачать отчет (.xlsx)",
            command=self._download_report,
            state=tk.DISABLED
        )
        self.download_button.pack(side=tk.LEFT)
        
        # Основная область с вкладками
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Вкладка с диаграммой
        self.chart_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.chart_frame, text="Диаграмма")
        
        # Фрейм для диаграммы
        self.pie_frame = ttk.Frame(self.chart_frame)
        self.pie_frame.pack(fill=tk.BOTH, expand=True)
        
        # Вкладка с таблицей
        self.table_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.table_frame, text="Таблица")
        
        # Создаем таблицу
        self._create_table()
        
        # Вкладка со сравнением (изначально скрыта)
        self.comparison_frame = ttk.Frame(self.notebook)
        
        # Нижняя панель с журналом
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Метка журнала
        ttk.Label(bottom_frame, text="Журнал событий:").pack(anchor=tk.W)
        
        # Текстовое поле для журнала
        self.log_text = tk.Text(bottom_frame, height=5, wrap=tk.WORD)
        self.log_text.pack(fill=tk.X)
        self.log_text.config(state=tk.DISABLED)
    
    def _create_table(self):
        """
        Создание таблицы для отображения данных.
        """
        # Фрейм для таблицы ролей
        roles_frame = ttk.LabelFrame(self.table_frame, text="Роли по трудозатратам")
        roles_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем Treeview для таблицы ролей
        columns = ("department", "role", "count", "cost")
        self.roles_tree = ttk.Treeview(roles_frame, columns=columns, show="headings")
        
        # Настраиваем заголовки
        self.roles_tree.heading("department", text="Отдел")
        self.roles_tree.heading("role", text="Роль")
        self.roles_tree.heading("count", text="Количество")
        self.roles_tree.heading("cost", text="Трудозатраты (руб.)")
        
        # Настраиваем ширину столбцов
        self.roles_tree.column("department", width=150)
        self.roles_tree.column("role", width=150)
        self.roles_tree.column("count", width=100)
        self.roles_tree.column("cost", width=150)
        
        # Добавляем полосу прокрутки
        scrollbar = ttk.Scrollbar(roles_frame, orient=tk.VERTICAL, command=self.roles_tree.yview)
        self.roles_tree.configure(yscrollcommand=scrollbar.set)
        
        # Размещаем элементы
        self.roles_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Фрейм для сводки по отделам
        summary_frame = ttk.LabelFrame(self.table_frame, text="Сводка по отделам")
        summary_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Создаем Treeview для сводки
        columns = ("department", "count", "cost")
        self.summary_tree = ttk.Treeview(summary_frame, columns=columns, show="headings", height=3)
        
        # Настраиваем заголовки
        self.summary_tree.heading("department", text="Отдел")
        self.summary_tree.heading("count", text="Количество")
        self.summary_tree.heading("cost", text="Трудозатраты (руб.)")
        
        # Настраиваем ширину столбцов
        self.summary_tree.column("department", width=150)
        self.summary_tree.column("count", width=100)
        self.summary_tree.column("cost", width=150)
        
        # Размещаем элементы
        self.summary_tree.pack(fill=tk.X)
    
    def _upload_file(self):
        """
        Обработчик нажатия кнопки загрузки файла.
        """
        # Открываем диалог выбора файла
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel файлы", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        # Отключаем кнопки на время загрузки
        self._set_loading_state(True)
        
        # Запускаем загрузку в отдельном потоке
        threading.Thread(target=self._process_file, args=(file_path,)).start()
    
    def _process_file(self, file_path):
        """
        Обработка выбранного файла.
        
        Args:
            file_path (str): Путь к выбранному файлу
        """
        try:
            # Загружаем файл
            self.log(f"Загрузка файла: {os.path.basename(file_path)}")
            success = self.loader.load_file(file_path)
            
            if not success:
                self.log("Ошибка при загрузке файла")
                self._set_loading_state(False)
                return
            
            # Получаем данные
            self.current_data = self.loader.get_current_data()
            
            # Анализируем данные
            self.log("Анализ данных...")
            self.current_results = self.analyzer.analyze_data(self.current_data)
            
            # Обновляем интерфейс
            self.root.after(0, self._update_ui)
            
            # Проверяем, нужно ли сравнение
            if self.compare_var.get():
                self._load_previous_data()
            
        except Exception as e:
            self.log(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка при обработке файла:\n{str(e)}")
            self.root.after(0, lambda: self._set_loading_state(False))
    
    def _update_ui(self):
        """
        Обновление интерфейса после загрузки данных.
        """
        # Создаем диаграмму
        self.log("Создание диаграммы...")
        self.current_figure = self.visualizer.create_pie_chart(self.analyzer)
        
        # Отображаем диаграмму
        self._display_chart(self.current_figure, self.pie_frame)
        
        # Обновляем таблицу
        self._update_table()
        
        # Включаем кнопку скачивания отчета
        self.download_button.config(state=tk.NORMAL)
        
        # Включаем кнопки
        self._set_loading_state(False)
        
        self.log("Данные успешно загружены и проанализированы")
    
    def _display_chart(self, figure, frame):
        """
        Отображение диаграммы в указанном фрейме.
        
        Args:
            figure (matplotlib.figure.Figure): Объект фигуры
            frame (ttk.Frame): Фрейм для отображения
        """
        # Очищаем фрейм
        for widget in frame.winfo_children():
            widget.destroy()
        
        # Создаем холст для диаграммы
        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas.draw()
        
        # Размещаем холст
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _update_table(self):
        """
        Обновление таблицы с данными.
        """
        # Очищаем таблицу ролей
        for item in self.roles_tree.get_children():
            self.roles_tree.delete(item)
        
        # Получаем отсортированный список ролей
        sorted_roles = self.analyzer.get_sorted_role_costs()
        
        # Добавляем данные в таблицу ролей
        for _, row in sorted_roles.iterrows():
            self.roles_tree.insert(
                "",
                tk.END,
                values=(
                    row["Отдел"],
                    row["Роль"],
                    row["Количество"],
                    f"{row['Трудозатраты']:.2f}"
                )
            )
        
        # Очищаем таблицу сводки
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        
        # Получаем сводку по отделам
        dept_summary = self.analyzer.get_department_summary()
        
        # Добавляем данные в таблицу сводки
        for _, row in dept_summary.iterrows():
            self.summary_tree.insert(
                "",
                tk.END,
                values=(
                    row["Отдел"],
                    row["Количество"],
                    f"{row['Трудозатраты']:.2f}"
                )
            )
    
    def _toggle_comparison(self):
        """
        Обработчик изменения состояния флажка сравнения.
        """
        if self.current_data is not None:
            if self.compare_var.get():
                self._load_previous_data()
            else:
                # Удаляем вкладку сравнения, если она есть
                if self.notebook.index("end") > 2:
                    self.notebook.forget(2)
    
    def _load_previous_data(self):
        """
        Загрузка данных предыдущего месяца для сравнения.
        """
        # Отключаем кнопки на время загрузки
        self._set_loading_state(True)
        
        # Запускаем загрузку в отдельном потоке
        threading.Thread(target=self._process_previous_data).start()
    
    def _process_previous_data(self):
        """
        Обработка данных предыдущего месяца.
        """
        try:
            # Загружаем предыдущий файл
            self.log("Загрузка предыдущего файла...")
            success = self.loader.load_previous_file()
            
            if not success:
                self.log("Не удалось загрузить предыдущий файл")
                self.root.after(0, lambda: self._set_loading_state(False))
                return
            
            # Получаем предыдущие данные
            previous_data = self.loader.get_previous_data()
            
            # Создаем новый анализатор для предыдущих данных
            previous_analyzer = DataAnalyzer()
            previous_results = previous_analyzer.analyze_data(previous_data)
            
            # Создаем сравнительные диаграммы
            self.log("Создание сравнительных диаграмм...")
            self.comparison_figure = self.visualizer.create_comparison_charts(
                self.analyzer, previous_analyzer
            )
            
            # Создаем таблицу сравнения
            comparison_df = self.analyzer.compare_with_previous(
                self.current_results, previous_results
            )
            
            # Обновляем интерфейс
            self.root.after(0, lambda: self._update_comparison_ui(comparison_df))
            
        except Exception as e:
            self.log(f"Ошибка при сравнении: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка при сравнении данных:\n{str(e)}")
            self.root.after(0, lambda: self._set_loading_state(False))
    
    def _update_comparison_ui(self, comparison_df):
        """
        Обновление интерфейса сравнения.
        
        Args:
            comparison_df (pandas.DataFrame): DataFrame с результатами сравнения
        """
        # Проверяем, есть ли уже вкладка сравнения
        if self.notebook.index("end") <= 2:
            # Добавляем вкладку сравнения
            self.notebook.add(self.comparison_frame, text="Сравнение")
        
        # Очищаем фрейм сравнения
        for widget in self.comparison_frame.winfo_children():
            widget.destroy()
        
        # Создаем фрейм для диаграмм
        charts_frame = ttk.Frame(self.comparison_frame)
        charts_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Отображаем сравнительные диаграммы
        self._display_chart(self.comparison_figure, charts_frame)
        
        # Создаем фрейм для таблицы сравнения
        table_frame = ttk.LabelFrame(self.comparison_frame, text="Сравнение по ролям")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем Treeview для таблицы сравнения
        columns = (
            "department", "role", 
            "current_count", "prev_count", "count_diff",
            "current_cost", "prev_cost", "cost_diff"
        )
        comparison_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        
        # Настраиваем заголовки
        comparison_tree.heading("department", text="Отдел")
        comparison_tree.heading("role", text="Роль")
        comparison_tree.heading("current_count", text="Текущее кол-во")
        comparison_tree.heading("prev_count", text="Пред. кол-во")
        comparison_tree.heading("count_diff", text="Разница кол-ва")
        comparison_tree.heading("current_cost", text="Текущие затраты")
        comparison_tree.heading("prev_cost", text="Пред. затраты")
        comparison_tree.heading("cost_diff", text="Разница затрат")
        
        # Настраиваем ширину столбцов
        comparison_tree.column("department", width=100)
        comparison_tree.column("role", width=100)
        comparison_tree.column("current_count", width=100)
        comparison_tree.column("prev_count", width=100)
        comparison_tree.column("count_diff", width=100)
        comparison_tree.column("current_cost", width=100)
        comparison_tree.column("prev_cost", width=100)
        comparison_tree.column("cost_diff", width=100)
        
        # Добавляем полосу прокрутки
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=comparison_tree.yview)
        comparison_tree.configure(yscrollcommand=scrollbar.set)
        
        # Размещаем элементы
        comparison_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Добавляем данные в таблицу сравнения
        for _, row in comparison_df.iterrows():
            comparison_tree.insert(
                "",
                tk.END,
                values=(
                    row["Отдел"],
                    row["Роль"],
                    row["Текущее количество"],
                    row["Предыдущее количество"],
                    row["Разница количества"],
                    f"{row['Текущие трудозатраты']:.2f}",
                    f"{row['Предыдущие трудозатраты']:.2f}",
                    f"{row['Разница трудозатрат']:.2f}"
                )
            )
        
        # Переключаемся на вкладку сравнения
        self.notebook.select(2)
        
        # Включаем кнопки
        self._set_loading_state(False)
        
        self.log("Сравнение с предыдущим месяцем выполнено")
    
    def _download_report(self):
        """
        Обработчик нажатия кнопки скачивания отчета.
        """
        if self.current_data is None or self.current_figure is None:
            messagebox.showwarning("Предупреждение", "Нет данных для создания отчета")
            return
        
        # Открываем диалог сохранения файла
        file_path = filedialog.asksaveasfilename(
            title="Сохранить отчет",
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        # Отключаем кнопки на время создания отчета
        self._set_loading_state(True)
        
        # Запускаем создание отчета в отдельном потоке
        threading.Thread(target=self._create_report, args=(file_path,)).start()
    
    def _create_report(self, file_path):
        """
        Создание отчета Excel.
        
        Args:
            file_path (str): Путь для сохранения отчета
        """
        try:
            # Создаем отчет
            self.log("Создание отчета Excel...")
            self.visualizer.create_excel_report(
                self.current_data,
                self.analyzer,
                self.current_figure,
                file_path
            )
            
            self.log(f"Отчет сохранен: {os.path.basename(file_path)}")
            self.root.after(0, lambda: messagebox.showinfo("Успех", f"Отчет успешно сохранен:\n{file_path}"))
            
        except Exception as e:
            self.log(f"Ошибка при создании отчета: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Произошла ошибка при создании отчета:\n{str(e)}"))
            
        finally:
            # Включаем кнопки
            self.root.after(0, lambda: self._set_loading_state(False))
    
    def _set_loading_state(self, is_loading):
        """
        Установка состояния загрузки для элементов интерфейса.
        
        Args:
            is_loading (bool): True, если идет загрузка, иначе False
        """
        state = tk.DISABLED if is_loading else tk.NORMAL
        self.upload_button.config(state=state)
        self.compare_checkbox.config(state=state)
        
        if self.current_data is not None:
            self.download_button.config(state=state)
    
    def log(self, message):
        """
        Добавление сообщения в журнал.
        
        Args:
            message (str): Текст сообщения
        """
        # Получаем текущее время
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Форматируем сообщение
        log_message = f"[{timestamp}] {message}\n"
        
        # Добавляем сообщение в журнал
        self.root.after(0, lambda: self._append_to_log(log_message))
    
    def _append_to_log(self, message):
        """
        Добавление сообщения в текстовое поле журнала.
        
        Args:
            message (str): Текст сообщения
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = NonconformanceAnalyzerApp(root)
    root.mainloop()
