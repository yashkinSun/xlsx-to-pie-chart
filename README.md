# Production Non-Conformance Analyzer

An application for analyzing production non-conformance data with visualization of results through multi-level pie charts and tables.

## Description

The application allows you to:

- Load non-conformance data from Excel files
- Analyze data by roles and departments (Production and Office)
- Visualize results using multi-level pie charts
- Display a role table sorted by labor costs
- Compare data with the previous month
- Generate Excel reports with source data, summary, and charts

## Requirements

- Python 3.8 or higher
- Libraries:
  - `tkinter` (for GUI)
  - `pandas` (for data processing)
  - `matplotlib` (for visualization)
  - `openpyxl` (for Excel operations)

## Installation

1. Install Python 3.8 or higher from the official website: https://www.python.org/downloads/

2. Install the required libraries:
   ```bash
   pip install pandas matplotlib openpyxl
   ```

3. Tkinter is usually included with Python for Windows. If not available, install it:
   ```bash
   pip install tk
   ```

## Project Structure

- `main.py` - Application entry point
- `app.py` - Main application class with GUI
- `data_loader.py` - Module for loading data from Excel
- `data_analyzer.py` - Module for data analysis
- `data_visualizer.py` - Module for data visualization
- `History/` - Folder for storing uploaded file history

## Running the Application

```bash
python main.py
```

## Usage

1. Click the "Load Excel File" button and select a file with non-conformance data
2. After loading the file, a pie chart and table with analysis results will be displayed
3. To compare with the previous month, check the "Compare with Previous Month" checkbox
4. To generate a report, click the "Download Report (.xlsx)" button

## Data Format

The application expects an Excel file with the following columns:

- `"Виновник (Производство )"` - production roles
- `"Виновник (Офис )"` - office roles
- `"Трудозатраты (рублей)"` - labor costs in rubles

Roles can be separated by "/" (e.g., "Programmer/Designer"). In this case, labor costs will be divided equally between roles.

## Creating an Executable File (.exe)

To create an executable file for Windows, use PyInstaller:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Create the executable file:
   ```bash
   pyinstaller --onefile --windowed main.py
   ```

3. The executable file will be created in the `dist/` folder

Alternatively, you can use Visual Studio 2022 to create an executable file:

1. Install Visual Studio 2022 with Python support
2. Open the project in Visual Studio
3. Select "Project" -> "Publish" menu and follow the wizard instructions

## Testing

To verify the application's functionality, run the test script:

```bash
python test_app.py
```

## Notes

- All uploaded files are automatically saved to the "History" folder with a timestamp
- When comparing with the previous month, the most recent file from the "History" folder is used
- The Excel report contains three sheets: source data, summary, and chart

---

**Beta Version Notice**: This application is currently in beta. Some features may be limited or subject to change.

**NOTE** Application main language is Russian. 