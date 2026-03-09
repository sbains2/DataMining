import pandas as pd
import numpy as np

def preprocess_for_all_models():
    input_file = "Car_Crash_Preprocessed.xlsx"
    output_file = "Car_Crash_Fully_Preprocessed.csv"
    
    print(f"Loading {input_file}...")
    df = pd.read_excel(input_file, engine='openpyxl')
    
    print("Initial Shape:", df.shape)
    
    # 1. Convert boolean columns to integer (0/1)
    bool_cols = df.select_dtypes(include=['bool']).columns
    print(f"Converting {len(bool_cols)} boolean columns to integers.")
    df[bool_cols] = df[bool_cols].astype(int)
    
    # 2. One-hot encode Categorical/Cyclical numeric columns
    cols_to_encode = ['County', 'Weekday', 'Month']
    print(f"One-hot encoding integer categorical columns: {cols_to_encode}")
    
    # Convert to string first so get_dummies treats them as categories
    for col in cols_to_encode:
        df[col] = df[col].astype(str)
        
    df = pd.get_dummies(df, columns=cols_to_encode, drop_first=False)
    
    # get_dummies generates True/False cols. Convert those to integer (0/1) as well
    new_bool_cols = df.select_dtypes(include=['bool']).columns
    if len(new_bool_cols) > 0:
        df[new_bool_cols] = df[new_bool_cols].astype(int)
    
    print("Final Shape after preprocessing:", df.shape)
    print("Final columns data types:")
    print(df.dtypes.value_counts())
    
    print(f"Saving to {output_file}...")
    df.to_csv(output_file, index=False)
    print("Done!")

if __name__ == "__main__":
    preprocess_for_all_models()
