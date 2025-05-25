import pandas as pd

# 读取两个Excel文件
file1_path = r'D:\AppDate\Python\project\copy_excel_data\data01.xlsx'
file2_path = r'D:\AppDate\Python\project\copy_excel_data\data02.xlsx'

# 加载Excel表格
df1 = pd.read_excel(file1_path)  # 第一张表，被提取
df2 = pd.read_excel(file2_path)  # 第二张表，按第二张表提取相应的列

# 获取第二张表的列名（假设第二张表的第一行是列名）
columns_to_extract = df2.columns.tolist()  # 获取列名

# 获取第二张表第一列的数据，用于提取对应的行
matching_values = df2.iloc[:, 0].tolist()

# 打印 df1 和 df2 的列名，检查它们是否匹配
print("df1 columns:", df1.columns.tolist())
print("columns_to_extract:", columns_to_extract)
print("匹配的值:", matching_values)

# 只保留在 df1 中实际存在的列
columns_to_extract_existing = [col for col in columns_to_extract if col in df1.columns]

# 检查哪些列在 df1 中没有找到，并打印警告信息
missing_cols = set(columns_to_extract) - set(columns_to_extract_existing)
if missing_cols:
    print(f"警告: 以下列在 df1 中找不到，将被跳过: {missing_cols}")

# 提取第一张表中存在的列
df_extracted = df1[columns_to_extract_existing]

# 根据第二张表的第一列匹配第一张表的行
df_extracted_matched = df_extracted[df_extracted.iloc[:, 0].isin(matching_values)]

# 按照第二张表第一列的顺序重新排序提取的数据
df_extracted_matched = df_extracted_matched.set_index(df_extracted_matched.columns[0])  # 设置索引为第一列
df_extracted_matched = df_extracted_matched.loc[matching_values]  # 按照匹配的值的顺序排序
df_extracted_matched = df_extracted_matched.reset_index()  # 重置索引

# 输出提取的数据
print("提取的数据（匹配的行，按顺序排列）：")
print(df_extracted_matched)

# 如果需要将结果保存到新的Excel文件
output_path = r'D:\AppDate\Python\project\copy_excel_data\data03.xlsx'
df_extracted_matched.to_excel(output_path, index=False)
