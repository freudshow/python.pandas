import pandas as pd
 
# 读取Excel文件
df = pd.read_excel('data.xlsx')
print(df.head())
print("--------------------------------------------------")
print(df)
