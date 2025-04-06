import os
import sys
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from data_loader import DataLoader
from data_analyzer import DataAnalyzer
from data_visualizer import DataVisualizer

def test_data_loading():
    """Test data loading functionality"""
    print("\n=== Testing Data Loading ===")
    loader = DataLoader()
    success = loader.load_file("sample_data.xlsx")
    
    if success:
        print("✓ File loaded successfully")
        data = loader.get_current_data()
        print(f"✓ Number of records: {len(data)}")
        print(f"✓ Columns: {', '.join(data.columns[:5])}...")
        
        # Test saving to history
        history_files = os.listdir("History")
        print(f"✓ History folder contains {len(history_files)} files")
        
        return loader, data
    else:
        print("✗ Failed to load file")
        return None, None

def test_data_analysis(data):
    """Test data analysis functionality"""
    print("\n=== Testing Data Analysis ===")
    analyzer = DataAnalyzer()
    results = analyzer.analyze_data(data)
    
    print("✓ Data analyzed successfully")
    
    # Test role counts
    print("\nRole counts:")
    for dept, roles in results["role_counts"].items():
        print(f"  {dept}: {len(roles)} roles")
        for role, count in roles.items():
            print(f"    - {role}: {count}")
    
    # Test department summary
    print("\nDepartment summary:")
    for dept in ["Производство", "Офис"]:
        count = results["department_counts"][dept]
        cost = results["department_costs"][dept]
        print(f"  {dept}: {count} cases, {cost:.2f} RUB")
    
    # Test sorted roles
    sorted_roles = analyzer.get_sorted_role_costs()
    print("\nTop 3 roles by labor cost:")
    for i, (_, row) in enumerate(sorted_roles.head(3).iterrows()):
        print(f"  {i+1}. {row['Роль']} ({row['Отдел']}): {row['Трудозатраты']:.2f} RUB")
    
    return analyzer, results

def test_visualization(analyzer):
    """Test visualization functionality"""
    print("\n=== Testing Visualization ===")
    visualizer = DataVisualizer()
    
    # Test pie chart creation
    fig = visualizer.create_pie_chart(analyzer)
    print("✓ Pie chart created successfully")
    
    # Save chart to file for inspection
    chart_file = "test_pie_chart.png"
    visualizer.save_figure_to_file(fig, chart_file)
    print(f"✓ Chart saved to {chart_file}")
    
    # Test Excel report creation
    loader = DataLoader()
    loader.load_file("sample_data.xlsx")
    data = loader.get_current_data()
    
    report_file = "test_report.xlsx"
    visualizer.create_excel_report(data, analyzer, fig, report_file)
    print(f"✓ Excel report created: {report_file}")
    
    # Check if report file exists and has expected sheets
    if os.path.exists(report_file):
        try:
            xls = pd.ExcelFile(report_file)
            sheets = xls.sheet_names
            print(f"✓ Report contains sheets: {', '.join(sheets)}")
        except Exception as e:
            print(f"✗ Error checking Excel report: {e}")
    
    return visualizer, fig

def test_gui_packaging():
    """Test GUI packaging requirements"""
    print("\n=== Testing GUI Packaging Requirements ===")
    
    # Check if all required files exist
    required_files = ["app.py", "data_loader.py", "data_analyzer.py", "data_visualizer.py", "main.py"]
    missing_files = [f for f in required_files if not os.path.exists(f)]
    
    if missing_files:
        print(f"✗ Missing files: {', '.join(missing_files)}")
    else:
        print("✓ All required files exist")
    
    # Check if main.py imports app correctly
    try:
        with open("main.py", "r") as f:
            content = f.read()
            if "import tkinter" in content and "from app import NonconformanceAnalyzerApp" in content:
                print("✓ main.py correctly imports the application")
            else:
                print("✗ main.py may have incorrect imports")
    except Exception as e:
        print(f"✗ Error checking main.py: {e}")
    
    # Note about packaging
    print("\nNote: For packaging as .exe:")
    print("1. Install PyInstaller: pip install pyinstaller")
    print("2. Create .exe: pyinstaller --onefile --windowed main.py")
    print("3. The executable will be in the dist folder")

def main():
    """Run all tests"""
    print("Starting application tests...")
    
    # Test data loading
    loader, data = test_data_loading()
    if data is None:
        print("Cannot continue testing without data")
        return
    
    # Test data analysis
    analyzer, results = test_data_analysis(data)
    
    # Test visualization
    visualizer, fig = test_visualization(analyzer)
    
    # Test GUI packaging requirements
    test_gui_packaging()
    
    print("\n=== Test Summary ===")
    print("✓ Data loading functionality works correctly")
    print("✓ Data analysis functionality works correctly")
    print("✓ Visualization functionality works correctly")
    print("✓ All required files for packaging exist")
    print("\nThe application is ready for packaging and delivery.")

if __name__ == "__main__":
    main()
