# Practice
練習用程式碼 Project_Judgement.py

## 安裝必要套件

### 環境
終端機或power shell輸入以下，檢查是否有python
```bash
python --version
```
Python 3.11.9

### 在執行程式前，需安裝以下套件：
```bash
pip install selenium beautifulsoup4 pandas openpyxl
```
另外，下載適合你瀏覽器的 WebDriver (例如 ChromeDriver) 並確認其路徑。
[ChromeDriver 下載網址](https://sites.google.com/chromium.org/driver/)

### 程式說明
分別有三個function
1. crawl_judgment_system(keyword, max_pages)
2. save_to_csv(data, filename, download_path)
3. save_to_xlsx(data, filename, download_path)

### 主程式
* **需更改chromedriver.exe的路徑**
* 可更改搜尋關鍵字(keyword)
* 可更改爬蟲的頁數(max_pages)，目前預設為2
* 可更改保存檔案的路徑(download_path)，目前預設為C:\Users\User\Downloads，檔案格式為.xlsx，
* 可以更改成csv檔案類型，需取消註解此save_to_csv

```pyhon=
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
```

