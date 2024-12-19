import csv
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook


def crawl_judgment_system(keyword, max_pages):
    judgments = []

    try:
        # 打開網站
        driver = webdriver.Chrome(service=service)  # 使用Chrome瀏覽器，這裡也可以更改為其他瀏覽器
        driver.get(url)

        # 輸入關鍵字並送出查詢
        search_box = driver.find_element(By.ID, "txtKW")
        search_box.send_keys(keyword)
        search_box.send_keys(Keys.RETURN)

        # 等待iframe載入
        wait = WebDriverWait(driver, 20)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))

        # 切換到 iframe
        iframe = driver.find_element(By.TAG_NAME, "iframe")
        driver.switch_to.frame(iframe)

        # 爬取資料
        soup = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.find("table", {"id": "jud"})
        rows = table.find_all("tr")[1:]  # 排除表頭


        # 抓取多頁內容
        for page in range(1, max_pages + 1):
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table = soup.find("table", {"id": "jud"})

            if table:
                rows = table.find_all("tr")[1:]  # 排除表頭

                for row in rows:
                    cols = row.find_all("td")
                    # 確保列長度足夠
                    if len(cols) >= 4:
                        judgment = {
                            "序號": cols[0].text.strip(),
                            "裁判字號 （內容大小）": cols[1].text.strip(),
                            "裁判日期": cols[2].text.strip(),
                            "裁判案由": cols[3].text.strip(),
                        }
                        judgments.append(judgment)

            # 如果還有下一頁，則點擊「下一頁」按鈕
            if page < max_pages:
                try:
                    next_page_button = driver.find_element(By.ID, "hlNext")
                    next_page_button.click()
                    time.sleep(2)  # 等待頁面加載
                except Exception as e:
                    print(f"無法點擊下一頁: {e}")
                    break
                
    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()
    
    return judgments

def save_to_csv(data, filename, download_path):
    # 確保下載路徑存在
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    # 組合檔案完整路徑
    full_path = os.path.join(download_path, filename)
    keys = data[0].keys()  # 获取字典键
    with open(full_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=keys)
        writer.writeheader()
        for item in data:
            item['序號'] = item['序號'].replace('.', '')  # 去掉點
        writer.writerows(data)

def save_to_xlsx(data, filename, download_path):
    # 確保下載路徑存在
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    # 組合檔案完整路徑
    full_path = os.path.join(download_path, filename)
    
    # 創建 Excel 工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "Judgments"

    # 寫入表頭
    if data:
        headers = data[0].keys()
        ws.append(list(headers))

        # 寫入資料並去掉序號的點號
        for row in data:
            row['序號'] = row['序號'].replace('.', '')  # 去掉點
            ws.append(list(row.values()))
            
    # 保存檔案s
    wb.save(full_path)

if __name__ == "__main__":

    #司法院裁判書系統網址
    url = "https://judgment.judicial.gov.tw/FJUD/default.aspx" 

    #更改chromedriver.exe的路徑
    service = Service("D:/Microsoft VS Code/Project/chromedriver-win64/chromedriver.exe") 

    #輸入想搜尋的關鍵字
    keyword = "拆屋還地" 
    
    #結果，max_pages為要爬取的頁數
    result = crawl_judgment_system(keyword, max_pages = 2)

    #保存檔案
    download_path = os.path.expanduser("C:\\Users\\User\\Downloads")
    #save_to_csv(result, 'judgments.csv',download_path)
    save_to_xlsx(result, 'judgments.xlsx', download_path)
