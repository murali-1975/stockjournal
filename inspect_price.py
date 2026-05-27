import pandas as pd
try:
    df = pd.read_excel('Transformed_Tradebook.xlsx', sheet_name='Price_Update')
    print("Price_Update columns:")
    print(df.columns.tolist())
    print("First few rows:")
    print(df.head())
except Exception as e:
    print(e)
