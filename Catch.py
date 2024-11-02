from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os

def setup_driver():
    return webdriver.Edge()

def read_excel(file_path):
    return pd.read_excel(file_path)

def write_to_excel(file_path, df):
    if os.path.exists(file_path):
        existing_df = pd.read_excel(file_path, engine='openpyxl', header=None)
        updated_df = pd.concat([existing_df, df], ignore_index=True)
    else:
        updated_df = df
    updated_df.to_excel(file_path, index=False, header=False, engine='openpyxl')

def scrape_data(driver, url, input1_value, input2_value):
    driver.get(url)
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_xingming'))).send_keys(input1_value)
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 's_chaxunma'))).send_keys(input2_value)
    
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(),"查询")]'))).click()

    try:
        # 等待数据加载，同时检查是否出现弹出窗口
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'right_cell')))
    except:
        # 如果出现弹出窗口，重新加载页面
        print("检测到弹出窗口，重新加载页面...")
        driver.get(url)
        return None  # 返回空值表示没有抓取到数据

    elements = driver.find_elements(By.CLASS_NAME, 'right_cell')
    return [element.text for element in elements]

def main():
    input_file = r"C:\Users\yaoyi\OneDrive\桌面\Grade\22.xlsx"
    output_file = r"C:\Users\yaoyi\OneDrive\桌面\1010考试\8班.xlsx"
    url = "https://i0aew1vd.yichafen.com/qz/p7E8OWcMPt"

    driver = setup_driver()
    df = read_excel(input_file)
    output_data = []

    try:
        for index, row in df.iterrows():
            data1 = row[1]
            data2 = str(row[2])[-4:]
            try:
                output_values = scrape_data(driver, url, data1, data2)
                if output_values is None:
                    continue  # 如果没有抓取到数据，跳到下一个
                output_data.append([data1, data2] + output_values)
            except Exception as e:
                print(f"处理第 {index} 行时出错: {e}")
                driver.get(url)
                continue

    finally:
        driver.quit()

    if output_data:
        num_output_columns = len(output_data[0]) - 2
        columns = ['Input 1', 'Input 2'] + [f'Output {i + 1}' for i in range(num_output_columns)]
        output_df = pd.DataFrame(output_data, columns=columns)
        write_to_excel(output_file, output_df)

    input("按回车键退出...")

if __name__ == "__main__":
    main()
