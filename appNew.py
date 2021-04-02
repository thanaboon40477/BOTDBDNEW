from selenium.webdriver.common.keys import Keys
import time
from funcSelenium import *
import sys

while True:
    lst = []
    while True:
        yesno = input("To select a page, enter y or Y, \nif you want all pages, enter n or N : ").lower()
        if (yesno == str(yesno)) and (yesno == 'y' or yesno == 'n'):
            break
        print("\nPlease fill in correctly.!!!")

    if yesno == "y":
        while True:
            cookie = str(input("Please enter Cookie Token. : "))
            if cookie == '':
                print("\Please fill in correctly.!!!")
            elif cookie != '':
                keyword = str(input("Please enter Keyword. : "))
                if keyword == '':
                    print("\nPlease fill in correctly.!!!")
                elif keyword != '':
                    break

        while True:
            try:
                startPage = int(input("Please Enter the Start page. : "))
                finishPage = int(input("Please Enter the Ending page. : "))
                start = startPage
                if (startPage == 0) and (finishPage == 0) :
                    print("\nPlease enter a number other than 0.")
                elif (finishPage < startPage):
                    print("\nPlease complete the ending page rather than the starting page.")
                elif (finishPage > startPage or finishPage == startPage):
                    break
            except ValueError:
                print("\nPlease enter a number")

        
        driver = open_browser()
        driver.get("https://datawarehouse.dbd.go.th/")
        enter_wesbite(driver, cookie, keyword)

        sum_page = finishPage - startPage

        input_page = driver.find_element_by_xpath('//*[@id="cPage"]')
        input_page.send_keys(Keys.CONTROL,'a')
        input_page.send_keys(Keys.BACK_SPACE)
        input_page.send_keys(startPage)
        input_page.send_keys(Keys.ENTER)
        time.sleep(8)
        input_page.send_keys(Keys.PAGE_UP)

        
        try:
            for i in range(sum_page+1):
                print()
                print('************** Page',startPage,'**************')
                box = scaping_row_table(driver, lst, startPage)
                next_page = driver.find_element_by_xpath('//*[@id="next"]')
                next_page.click()
                startPage += 1
                time.sleep(8)
            finish = startPage - 1
            save_excel_page_new(box, keyword, start, finish, driver)
            print(f"Completed Page {start} to Page {finish} :)\n")   
        except:
            save_excel_page_new(box, keyword, start, startPage-1, driver)
            print(f"Completed Page {start} to Page {startPage-1} :)\n")
                
    elif yesno == "n":
        cookie = str(input("Enter Cookie : "))
        keyword = str(input("Enter Keyword : "))

        driver = open_browser()
        driver.get("https://datawarehouse.dbd.go.th/")
        enter_wesbite(driver, cookie, keyword)

        try:
            startAll = 1
            while True:
                print('Page',i+1,'...')
                box = scaping_row_table(driver, lst, startAll)
                next_page = driver.find_element_by_xpath('//*[@id="next"]')
                next_page.click()
                time.sleep(8)
                startAll += 1
        except:
            save_excel_all_new(box, keyword, driver)
            print(f"Completed All Page :)\n")

    transaction = input("Will continue the transaction, \nenter y or Y, if no continue, enter n or N. : ").lower()
    print()
    while True:
        if transaction != 'y' and transaction != 'n':
            transaction = input("\nPlease enter (y or Y) or (n or N) : ")
        elif transaction == 'n' or transaction == 'y':
                break
    if transaction == 'n':
        sys.exit("The Program is Terminated....")

    
      



