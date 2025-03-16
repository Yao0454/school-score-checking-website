import sys
import os
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QCheckBox, QPushButton, QTextEdit, QFileDialog, QGroupBox, QFormLayout, QScrollArea, QSpinBox
from PyQt6.QtCore import Qt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []  # 用于存储选择的文件路径
        self.df = None  # 初始化 self.df
        self.initUI()

    def initUI(self):
        # 主布局
        main_layout = QVBoxLayout()

        # 输入文件路径选择
        input_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit(self)
        input_button = QPushButton("选择输入文件夹", self)
        input_button.clicked.connect(self.select_input_folder)  # 选择文件夹
        input_layout.addWidget(QLabel("输入文件夹:"))
        input_layout.addWidget(self.input_file_edit)
        input_layout.addWidget(input_button)

        # 创建文件显示区域
        self.file_list_groupbox = QGroupBox("Excel 文件列表")
        self.file_list_layout = QFormLayout()
        self.file_list_groupbox.setLayout(self.file_list_layout)

        # 可滚动的区域，并限制其大小
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.file_list_groupbox)
        scroll_area.setWidgetResizable(True)  # 让QScrollArea根据内容自动调整大小
        scroll_area.setFixedHeight(150)  # 限制显示区域的高度
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)  # 禁用水平滚动条

        # 输入文件夹选择
        main_layout.addLayout(input_layout)
        main_layout.addWidget(scroll_area)  # 将文件列表滚动区域嵌入布局中

        # 输出文件路径
        output_layout = QHBoxLayout()
        self.output_file_edit = QLineEdit(self)
        output_button = QPushButton("选择输出路径", self)
        output_button.clicked.connect(self.select_output_path)
        output_layout.addWidget(QLabel("输出文件路径:"))
        output_layout.addWidget(self.output_file_edit)
        output_layout.addWidget(output_button)

        # 自动更改文件名勾选框
        self.auto_name_checkbox = QCheckBox("自动更改文件名", self)
        self.auto_name_checkbox.toggled.connect(self.auto_generate_filename)

        # 手动修改文件名文本框
        self.manual_filename_edit = QLineEdit(self)
        self.manual_filename_edit.setPlaceholderText("手动输入文件名（如果没有勾选自动更改）")

        # URL 输入框
        self.url_edit = QLineEdit(self)
        self.url_edit.setPlaceholderText("请输入网址")

        # 线程数量调节
        thread_layout = QHBoxLayout()
        self.thread_spinbox = QSpinBox(self)
        self.thread_spinbox.setRange(1, 20)  # 设置线程数量范围
        self.thread_spinbox.setValue(5)  # 设置默认值
        thread_layout.addWidget(QLabel("线程数量:"))
        thread_layout.addWidget(self.thread_spinbox)

        # 运行按钮
        run_button = QPushButton("运行", self)
        run_button.clicked.connect(self.run_process)

        # 控制台输出
        self.console_output = QTextEdit(self)
        self.console_output.setReadOnly(True)

        # 添加布局
        main_layout.addLayout(output_layout)
        main_layout.addWidget(self.manual_filename_edit)
        main_layout.addWidget(self.auto_name_checkbox)
        main_layout.addWidget(self.url_edit)
        main_layout.addLayout(thread_layout)
        main_layout.addWidget(run_button)
        main_layout.addWidget(self.console_output)

        # 设置窗口属性
        self.setLayout(main_layout)
        self.setWindowTitle('数据抓取工具')
        self.setGeometry(100, 100, 600, 500)
        self.show()

    def select_input_folder(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输入文件夹")
        if folder_name:
            self.input_file_edit.setText(folder_name)
            self.load_files_in_folder(folder_name)

    def load_files_in_folder(self, folder_name):
        # 清空当前显示的文件列表
        for i in reversed(range(self.file_list_layout.count())):
            widget = self.file_list_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        # 获取文件夹中的所有 Excel 文件
        excel_files = [f for f in os.listdir(folder_name) if f.endswith(('.xlsx', '.xls'))]
        self.checkboxes = []  # 用于存储所有勾选框

        for file in excel_files:
            checkbox = QCheckBox(file)
            self.checkboxes.append((checkbox, os.path.join(folder_name, file)))  # 存储勾选框和文件路径的元组
            self.file_list_layout.addRow(checkbox)  # 添加到文件列表显示区域

    def select_output_path(self):
        folder_name = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if folder_name:
            self.output_file_edit.setText(folder_name)

    def auto_generate_filename(self):
        if self.auto_name_checkbox.isChecked():
            self.console_output.append("自动生成文件名已启用。")
            self.manual_filename_edit.setDisabled(True)  # 禁用手动修改文件名框
        else:
            self.manual_filename_edit.setDisabled(False)  # 恢复手动修改文件名框

    def run_process(self):
        selected_files = [file for checkbox, file in self.checkboxes if checkbox.isChecked()]
        if not selected_files:
            self.console_output.append("没有选择任何文件.")
            return

        output_path = self.output_file_edit.text()
        if not output_path:  # 确保输出路径不为空
            self.console_output.append("没有选择输出路径.")
            return

        # 确保输出路径是一个有效的文件夹
        if not os.path.isdir(output_path):
            self.console_output.append(f"输出路径无效: {output_path}")
            return

        # 如果没有勾选自动更改文件名，则使用手动输入的文件名
        if not self.auto_name_checkbox.isChecked():
            manual_filename = self.manual_filename_edit.text()
            if not manual_filename:
                self.console_output.append("没有输入手动文件名.")
                return
            output_file = os.path.join(output_path, manual_filename)
        else:
            # 自动生成文件名将在抓取完成后进行
            output_file = os.path.join(output_path, "temp_output.xlsx")

        # 确保生成的输出路径是有效的
        if not output_file.endswith(".xlsx"):
            output_file += ".xlsx"

        # 显示控制台输出
        self.console_output.append(f"输出文件路径: {output_file}")
        self.console_output.append("开始处理...")

        # 获取线程数量
        max_workers = self.thread_spinbox.value()

        # 并行处理选中的文件
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(self.process_data, input_file, output_file) for input_file in selected_files]
            for future in as_completed(futures):
                future.result()  # 等待每个任务完成

    def generate_auto_filename(self, data):
        # 根据需要生成自动文件名，这里使用抓取数据中的第二列第一行的值
        try:
            if data:
                new_filename = str(data[0][2]) + "班.xlsx"  # 使用抓取数据的第二列的值作为文件名
                return new_filename
        except Exception as e:
            self.console_output.append(f"生成自动文件名时出错: {e}")
        return "output.xlsx"

    def process_data(self, input_file, output_file):
        # 读取 Excel 文件
        df = pd.read_excel(input_file)
        self.console_output.append(f"读取输入文件: {input_file}")

        # 初始化 WebDriver（假设使用 Edge）
        try:
            driver = webdriver.Edge()
        except Exception as e:
            self.console_output.append(f"WebDriver 错误: {e}")
            return

        driver.get(self.url_edit.text())  # 从UI获取网址

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
                self.console_output.append(f"Error processing row {index}: {e}")
                driver.get(self.url_edit.text())  # 错误时重新加载页面
                continue

            driver.get(self.url_edit.text())  # 完成后返回原始页面

        # 在数据抓取完成后生成文件名（第二列第一行的值）
        if output_data:
            if self.auto_name_checkbox.isChecked():
                new_filename = self.generate_auto_filename(output_data)
                output_file = os.path.join(os.path.dirname(output_file), new_filename)
            self.console_output.append(f"新的输出文件名: {output_file}")

            # 保存输出数据
            output_df = pd.DataFrame(output_data)
            output_df.to_excel(output_file, index=False)
            self.console_output.append(f"输出文件已保存: {output_file}")

        driver.quit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec())