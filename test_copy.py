import pandas as pd
df1 = pd.DataFrame({'A': [1, 2, 3]})
df2 = df1.copy()
df2.at[0, 'A'] = '=FORMULA'
print("DF1:")
print(df1)
print("DF2:")
print(df2)
