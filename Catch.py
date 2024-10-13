from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os

# 读取Excel表格
df = pd.read_excel(r"C:\Users\yaoyi\OneDrive\桌面\Grade\22.xlsx")

# 初始化WebDriver（Edge浏览器）
driver = webdriver.Edge()

# 初始网址
url = "https://i0aew1vd.yichafen.com/qz/p7E8OWcMPt"
driver.get(url)

# 存储抓取到的数据
output_data = []

# 循环读取每一行的数据
for index, row in df.iterrows():
    data1 = row[1]  # 从 Excel 表格中读取第二列的值
    data2 = str(row[2])[-4:]  # 从 Excel 表格中读取第三列的后四位
    
    try:
        # 使用显式等待定位输入框，超时设为3秒
        input1 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_xingming')))
        input2 = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_chaxunma')))
        
        # 清除输入框中的内容
        input1.clear()
        input2.clear()
        
        # 将数据输入到输入框
        input1.send_keys(str(data1))
        input2.send_keys(str(data2))
        
        # 定位查询按钮并点击
        submit_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"查询")]')))
        submit_button.click()
        
        # 等待页面加载新的数据
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'right_cell')))
        
        # 抓取输出并保存
        elements = driver.find_elements(By.CLASS_NAME, 'right_cell')
        output_values = [element.text for element in elements]
        output_data.append([data1, data2] + output_values)
        
    except Exception as e:
        print(f"Error processing row {index}: {e}")
        # 返回初始网址
        driver.get(url)
        continue  # 跳过当前循环
    
    # 返回初始网址
    driver.get(url)

# 检查输出数据的列数量并动态生成列名
if output_data:
    num_output_columns = len(output_data[0]) - 2
else:
    num_output_columns = 0

columns = ['Input 1', 'Input 2'] + [f'Output {i+1}' for i in range(num_output_columns)]

# 将 output_data 转换为 DataFrame
output_df = pd.DataFrame(output_data, columns=columns)

# 将输出数据保存到 Excel
file_path = r"C:\Users\yaoyi\OneDrive\桌面\1010考试\8班.xlsx"

# 如果文件已存在，尝试读取现有数据并更新
if os.path.exists(file_path):
    try:
        existing_df = pd.read_excel(file_path, engine='openpyxl', header=None)
        updated_df = pd.concat([existing_df, output_df], ignore_index=True)
    except ValueError:
        updated_df = output_df
else:
    updated_df = output_df

# 保存更新后的 DataFrame 到 Excel 文件，不包含列头
updated_df.to_excel(file_path, index=False, header=False, engine='openpyxl')

# 保持浏览器窗口打开，等待用户手动关闭
input("Press Enter to close the browser and exit...")

# 关闭浏览器
driver.quit()
