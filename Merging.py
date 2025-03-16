import pandas as pd
import os

# 假设所有文件都在 'files' 目录下
folder_path = "C://Users//yaoyi//OneDrive//桌面//0313考试"

# 获取所有 Excel 文件的路径
file_paths = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.xlsx')]

# 初始化一个空的列表，用于存储所有文件的数据
all_data = []

# 遍历每个文件
for file_path in file_paths:
    # 读取文件中的所有工作表
    sheet_names = pd.ExcelFile(file_path).sheet_names
    
    # 遍历每个工作表
    for sheet in sheet_names:
        # 读取当前工作表的数据
        df = pd.read_excel(file_path, sheet_name=sheet)
        
        # 将当前工作表的数据添加到 all_data 中
        all_data.append(df)

# 合并所有数据
# 只保留第一个工作表的表头，其他的会忽略表头
merged_data = pd.concat(all_data, ignore_index=True)

# 保存合并后的数据到新的 Excel 文件
merged_data.to_excel('C://Users//yaoyi//OneDrive//桌面//0313考试//merged_output.xlsx', index=False)
