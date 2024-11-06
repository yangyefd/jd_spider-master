import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def jd_search_selenium(keyword, target_keyword, text_widget, driver):
    # 打开搜索页面
    driver.get(f"https://search.jd.com/Search?keyword={keyword}&enc=utf-8")

    try:
        # 等待商品列表元素加载完成
        wait = WebDriverWait(driver, 50)
        product_list = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="gl-i-wrap"]')))
        
        result = 0
        for item in product_list:
            # 尝试使用 XPATH 选择器找到商品链接
            try:
                link = item.find_element(By.XPATH, './/div[@class="p-name p-name-type-2"]/a')
                product_url = link.get_attribute('href')
                
                print('product_url:', product_url)
                if target_keyword in product_url:
                    result = 1
                    break
                
                time.sleep(1)  # 适当休眠，避免被识别为爬虫
            except Exception as e:
                print(f"Error finding link: {e}")

        # 在Text组件中显示结果
        if result == 1:
            text_widget.insert(tk.END, f"商品关键词: {keyword}, SPU: {target_keyword}, 检测结果: 找到商品\n")
        else:
            text_widget.insert(tk.END, f"商品关键词: {keyword}, SPU: {target_keyword}, 检测结果: 未找到商品\n")
        
        return result
    except Exception as e:
        print(f"An error occurred: {e}")
        return []

def read_excel(file_path, column_title="商品关键词"):
    target_SPUID = "所属商品SPUID"
    # 读取Excel文件
    wb = load_workbook(file_path)
    sheet = wb.active

    # 找到关键词所在的列
    header_row = sheet[1]  # 假设第一行是标题行
    keyword_column_index = None
    SPUID_column_index = None
    for cell in header_row:
        if cell.value == column_title:
            keyword_column_index = cell.column
            break
    for cell in header_row:
        if cell.value == target_SPUID:
            SPUID_column_index = cell.column
            break

    if keyword_column_index is None or SPUID_column_index is None:
        raise ValueError(f"Column with title '{column_title}' not found.")

    keywords = []
    SPUIDs = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # 根据找到的列索引读取关键词
        keyword = row[header_row.index(next(cell for cell in header_row if cell.column == keyword_column_index))]
        SPUID = row[header_row.index(next(cell for cell in header_row if cell.column == SPUID_column_index))]
        keywords.append(keyword)
        SPUIDs.append(SPUID)

    return keywords, SPUIDs

def search_from_excel(text_widget, driver):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    keywords, SPUIDs = read_excel(file_path)
    for keyword, target_keyword in zip(keywords, SPUIDs):
        jd_search_selenium(keyword, target_keyword, text_widget, driver)

# 创建Tkinter窗口
root = tk.Tk()
root.title("京东商品搜索")

# 创建Text组件用于显示信息
text_results = tk.Text(root, height=10, width=50)
text_results.pack()

# 创建浏览器实例
driver = webdriver.Edge()

# 创建按钮
button_search = tk.Button(root, text="选择Excel文件并搜索", command=lambda: search_from_excel(text_results, driver))
button_search.pack()

# 当窗口关闭时，关闭浏览器
root.protocol("WM_DELETE_WINDOW", lambda: [root.destroy(), driver.quit()])

# 运行Tkinter事件循环
root.mainloop()
