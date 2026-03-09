import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

f1 = "Big Data Files - Data Analytics.xlsx"
f2 = "Car_Crash_Preprocessed.xlsx"

print("--- Data Analytics original ---")
try:
    df1 = pd.read_excel(f1, engine='openpyxl')
    print("Shape:", df1.shape)
    print("Columns:", list(df1.columns))
    print("Null values per column:")
    print(df1.isnull().sum()[df1.isnull().sum() > 0])
    
    print("\nData Types:")
    print(df1.dtypes.value_counts())
    
except Exception as e:
    print("Error reading original:", e)

print("\n\n--- Car Crash Preprocessed ---")
try:
    df2 = pd.read_excel(f2, engine='openpyxl')
    print("Shape:", df2.shape)
    print("Columns:", list(df2.columns))
    print("Null values per column:")
    print(df2.isnull().sum()[df2.isnull().sum() > 0])
    
    print("\nData Types:")
    print(df2.dtypes.value_counts())
    
    print("\nSample of preprocessed data (first 3 rows):")
    print(df2.head(3))
except Exception as e:
    print("Error reading preprocessed:", e)

