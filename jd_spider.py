from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def jd_search_selenium(keyword, target_keyword):
    # 创建浏览器实例
    driver = webdriver.Chrome()
    # 打开搜索页面
    driver.get(f"https://search.jd.com/Search?keyword={keyword}&enc=utf-8")

    try:
        # 等待商品列表元素加载完成
        wait = WebDriverWait(driver, 50)
        product_list = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//div[@class="gl-i-wrap"]')))
        
        result_links = []
        for item in product_list:
            # 尝试使用 XPATH 选择器找到商品链接
            try:
                link = item.find_element(By.XPATH, './/div[@class="p-name p-name-type-2"]/a')
                product_url = link.get_attribute('href')
                
                # driver.execute_script("window.open(arguments[0]);", product_url)
                # driver.switch_to.window(driver.window_handles[1])
                
                # # 等待商品详情页加载完成
                # wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                
                # 检查页面内容中是否包含目标关键字
                if target_keyword in product_url:
                    result_links.append(product_url)
                
                # driver.close()
                # driver.switch_to.window(driver.window_handles[0])
                time.sleep(1)  # 适当休眠，避免被识别为爬虫
            except Exception as e:
                print(f"Error finding link: {e}")

        return result_links
    except Exception as e:
        print(f"An error occurred: {e}")
        return []
    finally:
        # 关闭浏览器
        driver.quit()

# 示例用法
keyword = "五妙水仙膏国药准字号"
target_keyword = "10097422656940"
result = jd_search_selenium(keyword, target_keyword)
if result:
    print("符合条件的链接地址：")
    for link in result:
        print(link)
else:
    print("未找到符合条件的商品")
