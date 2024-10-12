from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
import time

# 读取Excel表格
df = pd.read_excel(r"文件地址")

# 初始化WebDriver（Edge浏览器）
driver = webdriver.Edge()

# 初始网址
url = "https://i0aew1vd.yichafen.com/qz/p7E8OWcMPt"
driver.get(url)

# 等待页面加载
time.sleep(5)

# 存储抓取到的数据
output_data = []

# 循环读取每一行的数据
for index, row in df.iterrows():
    data1 = row[1]  # 从 Excel 表格中读取第二列的值
    data2 = str(row[2])[-4:]  # 从 Excel 表格中读取第三列的后四位
    
    # 定位并输入数据到网页的输入框中
    input1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 's_xingming')))  # 定位第一个输入框
    input2 = driver.find_element(By.NAME, 's_chaxunma')  # 定位第二个输入框

    input1.clear()  # 清除第一个输入框中的内容
    input2.clear()  # 清除第二个输入框中的内容
    
    input1.send_keys(str(data1))  # 将 data1 输入到第一个输入框
    input2.send_keys(str(data2))  # 将 data2 输入到第二个输入框

    # 定位带有“查询”字样的按钮并点击
    submit_button = driver.find_element(By.XPATH, '//button[contains(text(),"查询")]')
    submit_button.click()

    # 等待页面加载
    time.sleep(0.5)

    # 抓取输出并保存
    elements = driver.find_elements(By.CLASS_NAME, 'right_cell')  # 查找所有包含输出数据的元素
    output_values = [element.text for element in elements]  # 提取所有元素的文本
    output_data.append([data1, data2] + output_values)  # 将 data1, data2 和所有输出值作为一行数据添加到 output_data 中

    driver.get(url)  # 返回初始网址
    time.sleep(0.5)  # 等待页面加载

# 检查输出数据的列数量并动态生成列名
num_output_columns = len(output_data[0]) - 2  # 除去 'Input 1' 和 'Input 2' 列的输出列数
columns = ['Input 1', 'Input 2'] + [f'Output {i+1}' for i in range(num_output_columns)]

# 将 output_data 转换为 DataFrame
output_df = pd.DataFrame(output_data, columns=columns)

# 将输出数据保存到 Excel
file_path = '输出目录'

# 如果文件已存在，尝试读取现有数据并更新
if os.path.exists(file_path):
    try:
        # 读取现有数据，不使用列头
        existing_df = pd.read_excel(file_path, engine='openpyxl', header=None)
        # 将新数据追加到现有数据的末尾
        updated_df = pd.concat([existing_df, output_df], ignore_index=True)
    except ValueError:
        # 如果读取失败，直接使用新的 DataFrame
        updated_df = output_df
else:
    # 如果文件不存在，直接使用新的 DataFrame
    updated_df = output_df

# 保存更新后的 DataFrame 到 Excel 文件，不包含列头
updated_df.to_excel(file_path, index=False, header=False, engine='openpyxl')

# 保持浏览器窗口打开，等待用户手动关闭
input("Press Enter to close the browser and exit...")

# 关闭浏览器
driver.quit()
