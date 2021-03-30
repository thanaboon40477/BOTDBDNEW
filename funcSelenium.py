from selenium import webdriver
from bs4 import BeautifulSoup
import time
import os
import pandas as pd


def open_browser():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(os.path.abspath('chromedriver.exe'), options=options)
    return driver

def save_excel_page(lst, keyword, driver, start, finish):
    file = os.path.join(os.environ["USERPROFILE"], 'Desktop')
    file = os.path.join(file, f'Data_{keyword}_Page_{start}_to_{finish}.xlsx')
    cc = pd.concat(lst , ignore_index=True)
    excel = pd.ExcelWriter(file)
    cc.to_excel(excel, sheet_name="Sheet1", index=False)
    excel.save()
    driver.close()

def save_excel_all(lst, keyword, driver):
    file = os.path.join(os.environ["USERPROFILE"], 'Desktop')
    file = os.path.join(file, f'data_{keyword}.xlsx')
    cc = pd.concat(lst , ignore_index=True)
    excel = pd.ExcelWriter(file)
    cc.to_excel(excel, sheet_name="Sheet1", index=False)
    excel.save()
    driver.close()
    

def enter_wesbite(driver, cookie, keyword):
    driver.implicitly_wait(5)
    driver.delete_all_cookies() 
    driver.add_cookie({"name": "JSESSIONID", "domain": "datawarehouse.dbd.go.th",
                        "value": cookie})
    time.sleep(1)
    driver.get(f"https://datawarehouse.dbd.go.th/searchJuristicInfo/{keyword}/submitObjCode/1")

    span_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]')
    span_xpath.click()
    time.sleep(1)
    option_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]/option[5]')
    option_xpath.click()
    time.sleep(8)

def scaping_table(driver):
    soup = BeautifulSoup(driver.page_source, "lxml")
    page = soup.find_all('table', {"class":"table table-bordered table-striped table-hover text-center"})
    dfs = pd.read_html(str(page))
    data = pd.DataFrame(dfs[0])
    return data

# def scaping_row_table(driver, i):
#     soup = BeautifulSoup(driver.page_source, "lxml")
#     for i in range(11):
#         row_xpath = driver.find_element_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]')
#         row_xpath.click()
#         time.sleep(1)

#         Regis_number = soup.find('div', {'class':'row'}) # เลขทะเบียนนิติบุคคล
#         box = {"ลำดับ":}


  
