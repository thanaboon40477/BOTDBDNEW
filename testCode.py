from selenium import webdriver
import time
import os

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(os.path.abspath('chromedriver.exe'), options=options)
driver.get('https://www.settrade.com/C13_MarketSummary.jsp?detail=SET50')
time.sleep(4)

userid_element = driver.find_elements_by_xpath('/html/body/div[2]/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[3]/td[2]')
userid = userid_element
print(userid)