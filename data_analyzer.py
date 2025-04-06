import pandas as pd
import numpy as np

class DataAnalyzer:
    """
    Класс для анализа данных о несоответствиях.
    """
    
    def __init__(self):
        """
        Инициализация анализатора данных.
        """
        # Определение ролей по отделам
        self.production_roles = ["Резка", "Гибка", "Сборка", "Покраска", "Склад"]
        self.office_roles = ["Менеджер", "Расчётчик", "Конструктор", "Программист"]
        
        # Результаты анализа
        self.role_counts = None
        self.role_costs = None
        self.department_counts = None
        self.department_costs = None
        
    def analyze_data(self, df):
        """
        Анализ данных о несоответствиях.
        
        Args:
            df (pandas.DataFrame): Данные для анализа
            
        Returns:
            dict: Результаты анализа
        """
        if df is None or df.empty:
            raise ValueError("Нет данных для анализа")
        
        # Создаем копию данных для анализа
        data = df.copy()
        
        # Инициализируем словари для подсчета
        role_counts = {"Производство": {}, "Офис": {}}
        role_costs = {"Производство": {}, "Офис": {}}
        
        # Инициализируем счетчики для отделов
        department_counts = {"Производство": 0, "Офис": 0}
        department_costs = {"Производство": 0, "Офис": 0}
        
        # Обрабатываем каждую строку данных
        for _, row in data.iterrows():
            labor_cost = row["Трудозатраты (рублей)"]
            
            # Обработка ролей производства
            if pd.notna(row["Виновник (Производство )"]):
                roles = self._split_roles(row["Виновник (Производство )"])
                cost_per_role = labor_cost / len(roles)
                
                for role in roles:
                    role = role.strip()
                    if role not in role_counts["Производство"]:
                        role_counts["Производство"][role] = 0
                        role_costs["Производство"][role] = 0
                    
                    role_counts["Производство"][role] += 1
                    role_costs["Производство"][role] += cost_per_role
                    department_counts["Производство"] += 1
                    department_costs["Производство"] += cost_per_role
            
            # Обработка ролей офиса
            if pd.notna(row["Виновник (Офис )"]):
                roles = self._split_roles(row["Виновник (Офис )"])
                cost_per_role = labor_cost / len(roles)
                
                for role in roles:
                    role = role.strip()
                    if role not in role_counts["Офис"]:
                        role_counts["Офис"][role] = 0
                        role_costs["Офис"][role] = 0
                    
                    role_counts["Офис"][role] += 1
                    role_costs["Офис"][role] += cost_per_role
                    department_counts["Офис"] += 1
                    department_costs["Офис"] += cost_per_role
        
        # Сохраняем результаты
        self.role_counts = role_counts
        self.role_costs = role_costs
        self.department_counts = department_counts
        self.department_costs = department_costs
        
        # Возвращаем результаты в виде словаря
        return {
            "role_counts": role_counts,
            "role_costs": role_costs,
            "department_counts": department_counts,
            "department_costs": department_costs
        }
    
    def _split_roles(self, roles_str):
        """
        Разделение строки с несколькими ролями на отдельные роли.
        
        Args:
            roles_str (str): Строка с ролями, разделенными "/"
            
        Returns:
            list: Список отдельных ролей
        """
        if pd.isna(roles_str):
            return []
        
        # Разделяем строку по символу "/"
        roles = [r.strip() for r in roles_str.split("/")]
        return roles
    
    def get_role_data_for_chart(self):
        """
        Получение данных о ролях для построения графика.
        
        Returns:
            tuple: (labels, sizes, colors, department_labels)
        """
        if self.role_counts is None:
            raise ValueError("Данные не проанализированы")
        
        labels = []
        sizes = []
        colors = []
        department_labels = []
        
        # Цвета для отделов
        production_color = 'blue'
        office_color = 'orange'
        
        # Добавляем данные производства
        for role, count in self.role_counts["Производство"].items():
            labels.append(role)
            sizes.append(count)
            colors.append(production_color)
            department_labels.append("Производство")
        
        # Добавляем данные офиса
        for role, count in self.role_counts["Офис"].items():
            labels.append(role)
            sizes.append(count)
            colors.append(office_color)
            department_labels.append("Офис")
        
        return labels, sizes, colors, department_labels
    
    def get_sorted_role_costs(self):
        """
        Получение отсортированного списка ролей по трудозатратам.
        
        Returns:
            pandas.DataFrame: Отсортированный DataFrame с ролями и трудозатратами
        """
        if self.role_costs is None:
            raise ValueError("Данные не проанализированы")
        
        # Создаем список всех ролей и их трудозатрат
        roles_data = []
        
        for dept, costs in self.role_costs.items():
            for role, cost in costs.items():
                roles_data.append({
                    "Отдел": dept,
                    "Роль": role,
                    "Трудозатраты": cost,
                    "Количество": self.role_counts[dept][role]
                })
        
        # Создаем DataFrame и сортируем по трудозатратам
        df_roles = pd.DataFrame(roles_data)
        df_roles = df_roles.sort_values(by="Трудозатраты", ascending=False)
        
        return df_roles
    
    def get_department_summary(self):
        """
        Получение сводки по отделам.
        
        Returns:
            pandas.DataFrame: DataFrame со сводкой по отделам
        """
        if self.department_costs is None:
            raise ValueError("Данные не проанализированы")
        
        # Создаем список отделов и их показателей
        dept_data = []
        
        for dept in ["Производство", "Офис"]:
            dept_data.append({
                "Отдел": dept,
                "Количество": self.department_counts[dept],
                "Трудозатраты": self.department_costs[dept]
            })
        
        # Создаем DataFrame
        df_dept = pd.DataFrame(dept_data)
        
        return df_dept
    
    def compare_with_previous(self, current_results, previous_results):
        """
        Сравнение текущих результатов с предыдущими.
        
        Args:
            current_results (dict): Текущие результаты анализа
            previous_results (dict): Предыдущие результаты анализа
            
        Returns:
            pandas.DataFrame: DataFrame с результатами сравнения
        """
        if current_results is None or previous_results is None:
            raise ValueError("Нет данных для сравнения")
        
        # Создаем список всех ролей из обоих наборов данных
        all_roles = set()
        
        for dept in ["Производство", "Офис"]:
            all_roles.update(current_results["role_counts"][dept].keys())
            all_roles.update(previous_results["role_counts"][dept].keys())
        
        # Создаем данные для сравнения
        comparison_data = []
        
        for role in all_roles:
            # Определяем отдел для роли
            dept = "Производство" if role in self.production_roles else "Офис"
            
            # Получаем текущие значения
            current_count = current_results["role_counts"][dept].get(role, 0)
            current_cost = current_results["role_costs"][dept].get(role, 0)
            
            # Получаем предыдущие значения
            prev_count = previous_results["role_counts"][dept].get(role, 0)
            prev_cost = previous_results["role_costs"][dept].get(role, 0)
            
            # Вычисляем разницу
            count_diff = current_count - prev_count
            cost_diff = current_cost - prev_cost
            
            comparison_data.append({
                "Отдел": dept,
                "Роль": role,
                "Текущее количество": current_count,
                "Предыдущее количество": prev_count,
                "Разница количества": count_diff,
                "Текущие трудозатраты": current_cost,
                "Предыдущие трудозатраты": prev_cost,
                "Разница трудозатрат": cost_diff
            })
        
        # Создаем DataFrame и сортируем по разнице трудозатрат
        df_comparison = pd.DataFrame(comparison_data)
        df_comparison = df_comparison.sort_values(by="Разница трудозатрат", ascending=False)
        
        return df_comparison

# Тестирование класса
if __name__ == "__main__":
    from data_loader import DataLoader
    
    # Загружаем данные
    loader = DataLoader()
    loader.load_file("sample_data.xlsx")
    data = loader.get_current_data()
    
    # Анализируем данные
    analyzer = DataAnalyzer()
    results = analyzer.analyze_data(data)
    
    # Выводим результаты
    print("\nРоли по количеству:")
    for dept, roles in results["role_counts"].items():
        print(f"\n{dept}:")
        for role, count in roles.items():
            print(f"  {role}: {count}")
    
    print("\nРоли по трудозатратам:")
    for dept, roles in results["role_costs"].items():
        print(f"\n{dept}:")
        for role, cost in roles.items():
            print(f"  {role}: {cost:.2f} руб.")
    
    print("\nСводка по отделам:")
    for dept in ["Производство", "Офис"]:
        count = results["department_counts"][dept]
        cost = results["department_costs"][dept]
        print(f"  {dept}: {count} случаев, {cost:.2f} руб.")
    
    # Получаем отсортированный список ролей по трудозатратам
    sorted_roles = analyzer.get_sorted_role_costs()
    print("\nОтсортированный список ролей по трудозатратам:")
    print(sorted_roles)
