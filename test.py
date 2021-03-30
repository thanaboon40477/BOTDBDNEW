from selenium import webdriver
from bs4 import BeautifulSoup
import time
from lxml import html

# url = "https://www.bitkub.com/"
driver = webdriver.Chrome('chromedriver.exe')
driver.get("https://www.bitkub.com/")
driver.implicitly_wait(5)

time.sleep(2)
soup = BeautifulSoup(driver.page_source, "lxml")
driver.close()
print()