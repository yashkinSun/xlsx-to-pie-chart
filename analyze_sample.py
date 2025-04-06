import pandas as pd
import os

# Load the sample Excel file
file_path = 'sample_data.xlsx'
print(f"Loading file: {file_path}")

try:
    # Read the Excel file
    df = pd.read_excel(file_path)
    
    # Display basic information about the dataframe
    print("\nDataframe Info:")
    print(df.info())
    
    # Display the first few rows
    print("\nFirst 5 rows:")
    print(df.head())
    
    # List all column names
    print("\nColumn names:")
    for col in df.columns:
        print(f"- {col}")
    
    # Check if required columns exist
    required_columns = ["Трудозатраты", "Виновник (Производство )", "Виновник (Офис )"]
    print("\nChecking required columns:")
    for col in required_columns:
        if col in df.columns:
            print(f"- {col}: Found")
            # Show unique values for role columns
            if col != "Трудозатраты":
                unique_values = df[col].dropna().unique()
                print(f"  Unique values: {unique_values}")
        else:
            print(f"- {col}: Not found")
    
    # Check data types
    print("\nData types:")
    print(df.dtypes)
    
    # Check for missing values
    print("\nMissing values count:")
    print(df.isnull().sum())
    
except Exception as e:
    print(f"Error loading file: {e}")
