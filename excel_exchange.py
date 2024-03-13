import pandas as pd

# 读取Excel文件
df = pd.read_excel('out.xlsx')

# 将一列数据转换为一行数据
df = df.T

# 将第一行作为表头
df.columns = df.iloc[0]

# 删除第一行
df = df[1:]

# 保存到新的Excel文件
df.to_excel('new_file.xlsx', index=False)