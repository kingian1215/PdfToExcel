from PyPDF2 import PdfFileReader, PdfFileWriter
import tkinter as tk
from tkinter import filedialog
import os.path 


root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
# print(file_path)
file_name = os.path.basename(file_path).split('.')[0]
# print(file_name)
path = os.path.split(file_path)[0]
# print(path)
pdf_reader = PdfFileReader(file_path)
pagnum = 0
all_page = pdf_reader.getNumPages()
# print("all_page=", all_page)

if all_page % 10: 
    # 判斷是否有餘數 
    pagnum = all_page // 10 + 1 
    # 取餘數整數部分加1 
else: 
    pagnum = all_page // 10
# print("pagnum=", pagnum)

'''
for page in range(1, pagnum+1):
    # print("第"+str(page)+"份")
    pdf_writer = PdfFileWriter()
 
    for i in range((page-1)*10, page*10 if page != pagnum else all_page):
       # 遍歷到每一頁挨個生成寫入器
        pdf_writer.addPage(pdf_reader.getPage(i))
        # 寫入器被新增一頁後立即輸出產生pdf
        with open(path + '/{}-{}.pdf'.format(file_name, page), 'wb') as out:
            pdf_writer.write(out)
            # print("寫入第",i+1,"頁")
print("PDF分割完成")
'''
# 1.匯入selenium
from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import re

# 2.開啟瀏覽器

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(options=options)

# driver = webdriver.Chrome('chromedriver.exe')
url = "https://zhtw.109876543210.com/"

'''
driver.get('https://www.google.com.tw')
time.sleep(2000)
driver.close()
'''

for page in range(1, pagnum+1):
    # ocr_path = path + '/{}-{}.pdf'.format(file_name, page)
    ocr_path = 'D:\Python Work\PDF ORC\PDFtoEXCEL.pdf'
    # print(ocr_path)
    # 3.開啟頁面(頁面地址)
    
    driver.get(url)
    # driver.maximize_window()
    time.sleep(1000)
    # print("sleep")
    # 4.定位上傳檔案按鈕
    # upfile = driver.find_element_by_id("selectfiles")
    
    # 5.使用send_keys方法上傳檔案
    # upfile.send_keys(path + '\{}-{}.pdf'.format(file_name, page))
    # driver.find_elements_by_id("container").send_keys(ocr_path)
    driver.find_element_by_xpath("//div[@id='container'][@style='position: relative;']/input[id='selectfiles'][value='選擇文件']").send_keys(ocr_path)

    time.sleep(200)
    Select(driver.find_element_by_xpath("//div[@class='ocr_language']/select[@id='ocr_language_1']")).select_by_value('zhtw')
    # lang1 = driver.find_element_by_id('ocr_language_1')
    # Select(lang1).select_by_value('zhtw')
    # lang1.click()
    Select(driver.find_element_by_xpath("//div[@class='ocr_language']/select[id='ocr_language_2']")).select_by_value('zhcn')
    # lang2 = driver.find_element_by_id("ocr_language_2")
    # Select(lang1).select_by_value("en")
    # lang2.click()
    driver.find_element_by_xpath("//*[@id='free_convert']").click()
    # driver.find_element_by_xpath("//span[@id="ocr_language_2"]/input").click("英文")
    # driver.find_element_by_xpath("//span[@id="ocr_output_format_xlsx"]/input").click()
    # driver.find_element_by_xpath("//span[@id="free_convert"]/input").click()
    time.sleep(500)
    driver.find_element_by_class_name("an_niu_1").click()

    # 6.關閉瀏覽器
    driver.quit()