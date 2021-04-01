from selenium import webdriver
from bs4 import BeautifulSoup
import time
import os
import pandas as pd
from selenium.webdriver.common.keys import Keys

#-----------------------Scraping Table------------------------#
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

def scaping_table(driver):
    soup = BeautifulSoup(driver.page_source, "lxml")
    page = soup.find_all('table', {"class":"table table-bordered table-striped table-hover text-center"})
    dfs = pd.read_html(str(page))
    data = pd.DataFrame(dfs[0])
    return data
#-----------------------Scraping Table------------------------#


def open_browser():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(os.path.abspath('chromedriver.exe'), options=options)
    return driver
    
def enter_wesbite(driver, cookie, keyword):
    driver.implicitly_wait(5)
    driver.delete_all_cookies() 
    driver.add_cookie({"name": "JSESSIONID", "domain": "datawarehouse.dbd.go.th",
                        "value": cookie})
    time.sleep(1)
    driver.get(f"https://datawarehouse.dbd.go.th/searchJuristicInfo/{keyword}/submitObjCode/1")

    time.sleep(1)
    span_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]')
    span_xpath.click()
    time.sleep(1)
    option_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]/option[5]')
    option_xpath.click()
    time.sleep(8)

#-----------------------Scraping Location------------------------#
def scaping_row_table(driver, lst, startPage):
    for i in range(10):
        # start_time = time.time()
        i += 1
        print(f'------Page {startPage} Row',i,'...')
        # ลำดับ
        time.sleep(2)
        order = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[1]')
        number = order[0].text
        print(number)

        # เลขทะเบียนนิติบุคคล
        regis_number = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[2]')
        regis_nb = regis_number[0].text
        print(regis_nb)

        # ชื่อนิติบุคคล
        name_entity = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[3]')
        name_ett = name_entity[0].text
        print(name_ett)

        # ประเภทนิติบุคคล
        juristic_type = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[4]')
        jrt_type = juristic_type[0].text
        print(jrt_type)
        
        # สถานะ
        juristic_status = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[5]')
        jrt_status = juristic_status[0].text
        print(jrt_status)

        # รหัสประเภทธุรกิจ
        business_type_code = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[6]')
        bsn_type_code = business_type_code[0].text
        print(bsn_type_code)

        # ชื่อประเภทธุรกิจ
        name_business_type = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[7]')
        n_bsn_type = name_business_type[0].text
        print(n_bsn_type)

        # จังหวัด
        province = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[8]')
        prov = province[0].text
        print(prov)

        # ทุนจดทะเบียน
        registered_capital = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[9]')
        regis_capital = registered_capital[0].text
        print(regis_capital)

        # รายได้รวม
        total_income = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[10]')
        tt_income = total_income[0].text
        print(tt_income)

        # กำไร(ขาดทุน)
        profit_n_loss = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[11]')
        p_n_l = profit_n_loss[0].text
        print(p_n_l)

        # สินทรัพย์รวม
        total_assets = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[12]')
        tt_assets = total_assets[0].text
        print(tt_assets)

        # ส่วนของผู้ถือหุ้น
        equity = driver.find_elements_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]/td[13]')
        eqt = equity[0].text
        print(eqt)

        #---Click---#
        row_xpath = driver.find_element_by_xpath(f'//*[@id="fixTable"]/tbody/tr[{i}]')
        row_xpath.click()
        time.sleep(7)
        #---Click---#

        # วันที่จดทะเบียนจัดตั้ง
        date_incorporation = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/table/tbody/tr[4]/td[2]')                                                  
        date_incop = date_incorporation[0].text
        print(date_incop)

        # ที่ตั้ง
        the_location = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[2]')
        location = the_location[0].text
        print(location)

        # โทรศัพท์
        phone_number = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[3]/td[2]')
        phone_nb = phone_number[0].text
        print(phone_nb)

        # โทรสาร
        fax_number = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[4]/td[2]')
        fax = fax_number[0].text
        print(fax)

        # Website
        web_site = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[5]/td[2]')
        web = web_site[0].text
        print(web)

        # E-mail address
        email_adress = driver.find_elements_by_xpath('/html/body/div/div[4]/div[2]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[6]/td[2]')
        email = email_adress[0].text
        print(email)

        print(f"------Complate Page {startPage} Row {i} :)")
        print()

        # end_time = time.time() - start_time
        # print("Time : ", round(end_time, 2))
        
        driver.back() 
        select_val = driver.find_element_by_xpath('//*[@id="sortBy"]') 
        select_val.click()
        time.sleep(1)
        val_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]/option[1]')
        val_xpath.click()
        time.sleep(1)
        val_xpath = driver.find_element_by_xpath('//*[@id="sortBy"]/option[5]')
        val_xpath.click()
        time.sleep(1)
        input_page = driver.find_element_by_xpath('//*[@id="cPage"]')
        input_page.send_keys(Keys.BACK_SPACE)
        input_page.send_keys(startPage)
        time.sleep(1)
        input_page.send_keys(Keys.ENTER)
        time.sleep(8)

        data = {'ลำดับ':number, 'เลขทะเบียนนิติบุคคล':regis_nb, 'ชื่อนิติบุคคล':name_ett, 'ประเภทนิติบุคคล':jrt_type, 'สถานะ':jrt_status,
                    'รหัสประเภทธุรกิจ':bsn_type_code, 'ชื่อประเภทธุรกิจ':n_bsn_type, 'จังหวัด':prov, 'ทุนจดทะเบียน(บาท)':regis_capital,
                        'รายได้รวม(บาท)':tt_income, 'กำไร(ขาดทุน)สุทธิ(บาท)':p_n_l, 'สินทรัพย์รวม(บาท)':tt_assets, 'ส่วนของผู้ถือหุ้น(บาท)':eqt,
                            'วันที่จดทะเบียนจัดตั้ง':date_incop, 'ที่ตั้ง':location, 'โทรศัพท์':phone_nb, 'เว็บไซต์':web, 'อีเมลล์':email}
        
        lst.append(data)
        
    return lst
        
def save_excel_page_new(lst, keyword, start, finish, driver):
    file = os.path.join(os.environ["USERPROFILE"], 'Desktop')
    file = os.path.join(file, f'Data_{keyword}_Page_{start}_to_{finish}.xlsx')
    box = pd.DataFrame(lst)
    excel = pd.ExcelWriter(file)
    box.to_excel(excel, sheet_name="Sheet1", index=False)
    excel.save()
    driver.close()

def save_excel_all_new(lst, keyword, driver):
    file = os.path.join(os.environ["USERPROFILE"], 'Desktop')
    file = os.path.join(file, f'data_{keyword}.xlsx')
    box = pd.DataFrame(lst)
    excel = pd.ExcelWriter(file)
    box.to_excel(excel, sheet_name="Sheet1", index=False)
    excel.save()
    driver.close()
#-----------------------Scraping Location------------------------#



  
