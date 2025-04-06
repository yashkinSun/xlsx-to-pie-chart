import pandas as pd
import os
import datetime
import shutil
from pathlib import Path

class DataLoader:
    """
    Класс для загрузки и обработки данных из Excel-файлов с несоответствиями.
    """
    
    def __init__(self, history_folder="History"):
        """
        Инициализация загрузчика данных.
        
        Args:
            history_folder (str): Путь к папке для хранения истории файлов
        """
        self.history_folder = history_folder
        self.current_data = None
        self.previous_data = None
        
        # Создаем папку истории, если она не существует
        if not os.path.exists(history_folder):
            os.makedirs(history_folder)
    
    def load_file(self, file_path):
        """
        Загрузка данных из Excel-файла и сохранение копии в папку истории.
        
        Args:
            file_path (str): Путь к Excel-файлу
            
        Returns:
            bool: True, если загрузка прошла успешно, иначе False
        """
        try:
            # Загрузка данных из Excel
            df = pd.read_excel(file_path)
            
            # Проверка наличия необходимых столбцов
            required_columns = [
                "Виновник (Производство )", 
                "Виновник (Офис )", 
                "Трудозатраты (рублей)"
            ]
            
            # Проверяем наличие всех необходимых столбцов
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"В файле отсутствуют следующие столбцы: {', '.join(missing_columns)}")
            
            # Сохраняем данные
            self.current_data = df
            
            # Копируем файл в папку истории с временной меткой
            self._save_to_history(file_path)
            
            return True
            
        except Exception as e:
            print(f"Ошибка при загрузке файла: {e}")
            return False
    
    def _save_to_history(self, file_path):
        """
        Сохранение копии файла в папку истории с временной меткой.
        
        Args:
            file_path (str): Путь к исходному файлу
        """
        # Получаем имя файла без пути
        file_name = os.path.basename(file_path)
        
        # Добавляем временную метку к имени файла
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name, ext = os.path.splitext(file_name)
        history_file_name = f"{base_name}_{timestamp}{ext}"
        
        # Полный путь для сохранения
        history_path = os.path.join(self.history_folder, history_file_name)
        
        # Копируем файл
        shutil.copy2(file_path, history_path)
        print(f"Файл сохранен в истории: {history_path}")
    
    def load_previous_file(self):
        """
        Загрузка последнего файла из папки истории.
        
        Returns:
            bool: True, если загрузка прошла успешно, иначе False
        """
        try:
            # Получаем список файлов в папке истории
            history_files = [os.path.join(self.history_folder, f) for f in os.listdir(self.history_folder) 
                            if f.endswith('.xlsx') or f.endswith('.xls')]
            
            if not history_files:
                print("В папке истории нет файлов")
                return False
            
            # Сортируем файлы по времени изменения (от новых к старым)
            history_files.sort(key=os.path.getmtime, reverse=True)
            
            # Берем самый свежий файл (исключая текущий, если он есть)
            latest_file = history_files[0]
            
            # Загружаем данные
            self.previous_data = pd.read_excel(latest_file)
            print(f"Загружен предыдущий файл: {latest_file}")
            
            return True
            
        except Exception as e:
            print(f"Ошибка при загрузке предыдущего файла: {e}")
            return False
    
    def get_current_data(self):
        """
        Получение текущих данных.
        
        Returns:
            pandas.DataFrame: Текущие данные или None, если данные не загружены
        """
        return self.current_data
    
    def get_previous_data(self):
        """
        Получение предыдущих данных.
        
        Returns:
            pandas.DataFrame: Предыдущие данные или None, если данные не загружены
        """
        return self.previous_data

# Тестирование класса
if __name__ == "__main__":
    loader = DataLoader()
    success = loader.load_file("sample_data.xlsx")
    
    if success:
        print("Файл успешно загружен")
        data = loader.get_current_data()
        print(f"Количество записей: {len(data)}")
        print(f"Столбцы: {data.columns.tolist()}")
    else:
        print("Не удалось загрузить файл")
