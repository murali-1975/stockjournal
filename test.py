import pandas as pd
df = pd.read_excel('Transformed_Tradebook.xlsx', sheet_name='Transaction')
df['VN'] = df['Total_Value'].str.replace('₹', '', regex=False).str.replace(',', '', regex=False).astype(float)
df_sub = df[df['Symbol'].isin(['AVANTEL', 'CANBK', 'GMDCLTD', 'SPLPETRO', 'DMART']) & (df['Trade Type'] == 'buy')]
print(df_sub[['Trade Date', 'Symbol', 'VN', 'Tranches/Cheat']].to_string(index=False))
