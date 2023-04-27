import re
from selenium import webdriver
from fake_useragent import UserAgent
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException       
import xlsxwriter 
import pandas as pd
import numpy as np

# from symbol import parameters




useragent = UserAgent()

headers = {'Uset-Agent':
            f'{useragent.random}'
        }

def read_excel():
    data = pd.read_excel(r'C:\Users\koooo\Desktop\Белгородская областьF.СтроительствоКоммерческие корпоративные организации.xlsx', usecols='A')
    df = pd.DataFrame(data)
    links = df.values.tolist()
    return links

def check_exists_by_xpath(xpath, driver):
    try:
       driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

def create_Excel(parametr):
    try:
        book = xlsxwriter.Workbook(r'C:\Users\koooo\Desktop\xui5.xlsx')
        page = book.add_worksheet('')
        
        row = 0
        column = 0

        page.set_column('A:A', 50)
        page.set_column('B:B', 50)
        page.set_column('C:C', 50)
        page.set_column('D:D', 50)
        page.set_column('E:E', 50)
        page.set_column('F:F', 50)
        page.set_column('G:G', 50)
        page.set_column('H:H', 50)
        page.set_column('I:I', 50)

        for item in parametr():
            print(item)
            page.write(row, column, item[0])
            page.write(row, column+1, item[1])
            page.write(row, column+2, item[2])
            page.write(row, column+3, item[3])
            page.write(row, column+4, item[4])
            page.write(row, column+5, item[5])
            page.write(row, column+6, item[6])
            page.write(row, column+7, item[7])
            page.write(row, column+8, item[8])
            row += 1
        
    except:
        pass
    finally:
        book.close()

# Авторизация
def array():
    k = 0
    for url in read_excel():
        if k > 5:
            break
        for link in url:
            try:
                options = webdriver.FirefoxOptions()
                options.set_preference("general.useragent.override", useragent.random)
                options.set_preference('dom.webdriver.enabled', False)
                options.headless = True
                driver = webdriver.Firefox(
                    executable_path=r"C:\Users\koooo\Desktop\Поиск ссылок по Рф компаниям\firefox\geckodriver.exe",
                    options=options
                )

                driver.get(link)
                time.sleep(7)

                login = driver.find_element(By.XPATH, '/html/body/div[2]/header/div/div[3]/div')
                login.click()
                time.sleep(5)

                email_input = driver.find_element(By.ID, "mw-l_mail")
                email_input.clear()
                email_input.send_keys("perervads@mail.ru")
                time.sleep(2)

                password_input = driver.find_element(By.ID, "mw-l_pass")
                password_input.clear()
                password_input.send_keys('Kmg87055775544')
                time.sleep(2)

                password_input.send_keys(Keys.ENTER)
                time.sleep(5)
                telephone_input_list = []
                if check_exists_by_xpath('/html/body/div[5]/div/div/div/div/div[6]', driver) == True:
                    btn_activate = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div/div/div[6]')    
                    btn_activate.click()
                    time.sleep(7)
                if check_exists_by_xpath('/html/body/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div/span[4]/button', driver) == True:
                    telephone_input = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div/span[4]/button')
                    telephone_input_list = driver.find_elements(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div/div/span')
                    telephone_input.click()
                    time.sleep(2)

                name_company = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[1]/div[1]/div[1]').text
                OKD_company = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[1]/span[2]').text
                ogrn_company = driver.find_element(By.XPATH, '//*[@id="clip_ogrn"]').text
                inn_company = driver.find_element(By.XPATH, '//*[@id="clip_inn"]').text
                adres_company = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[1]/div[2]/address').text
                rukovoditel_company = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[1]/div[3]/span[3]/a/span').text
                nalog_company = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/span[2]').text
                telephone_company = driver.find_elements(By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div[1]/div[1]/div/div/div[2]/div[2]/div[4]/div[1]/div/span')
                revenue = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[1]/div[2]/div[2]/div[11]/div/div/div[2]/div[1]/div[2]').text

                all_telephone_list = telephone_company + telephone_input_list
                number = []
                number_telephone = []
                number_list = []
                number_sort = []
                for s in all_telephone_list:
                    s = s.text
                    only_number = re.sub('[^0-9]', '', s)
                    number.append(only_number)

                for q in range(len(number)):
                    if number[q] != '':
                        number_list.append(number[q])
                    else :
                        pass

                for c in number_list:

                    if len(c) == 11:
                        number_telephone.append(c)
                    elif len(c) == 12:
                        c = c[0:11]
                        number_telephone.append(c)
                    else:
                        pass

                for b in number_telephone:
                    b = b[:0] + '8' + b[0+1:]
                    number_sort.append(b)
                
                string_telephone = ''

                for h in number_sort:
                    string_telephone += str(h) + '  '
   
                info_list = []

                info_list.append(name_company)
                info_list.append(OKD_company)
                info_list.append(ogrn_company)
                info_list.append(inn_company)
                info_list.append(adres_company)
                info_list.append(rukovoditel_company)
                info_list.append(nalog_company)
                info_list.append(string_telephone)
                info_list.append(revenue)
                k += 1
                print(k)
                
                # print(info_list)
                yield info_list
            except:
                pass
            finally:
                driver.close()
                driver.quit()

create_Excel(array)