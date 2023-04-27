from selenium import webdriver
import requests as req
from bs4 import BeautifulSoup as BS
from fake_useragent import UserAgent
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException       
import xlsxwriter 

useragent = UserAgent()

headers = {'Uset-Agent':
            f'{useragent.random}'
        }
region_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[5]/div[2]/div[2]/div[2]/ul/li[7]/ul/li[1]/div/label'
activity_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[3]/div[2]/div[2]/div[2]/ul/li[6]/div'
legal_form_one_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[9]/div[2]/div[2]/div[2]/ul/li[1]/div'
legal_form_two_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[9]/div[2]/div[2]/div[2]/ul/li[2]/div'
legal_form_three_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[9]/div[2]/div[2]/div[2]/ul/li[3]/div'
legal_form_four_input_def = '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[9]/div[2]/div[2]/div[2]/ul/li[4]/div'


def check_exists_by_xpath(xpath, driver):
    try:
       driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

def create_Excel(list_links, region_excel, activity_excel, legal_form_excel):
    try:

        name = region_excel + activity_excel + legal_form_excel
        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\{name}.xlsx')
        page = book.add_worksheet('')
        

        row = 0
        column = 0

        page.set_column('A:A', 50)

        for item in list_links:
            page.write(row, column, item)
            row += 1
        
    except Exception as ex:
        print(ex)
    finally:
        book.close()
# Поиск ссылок 
def find_link(list_links, driver):
        company_items = driver.find_element(By.CLASS_NAME, 'additional-results')
        company_item_list = company_items.find_elements(By.TAG_NAME, 'a')
        for i in company_item_list:
            list_links.append(i.get_attribute('href'))
# Авторизация
def authentication():
    options = webdriver.FirefoxOptions()
    options.set_preference("general.useragent.override", useragent.random)
    options.set_preference('dom.webdriver.enabled', False)
    # options.headless = True

    driver = webdriver.Firefox(
        executable_path=r"C:\Users\koooo\Desktop\Поиск ссылок по Рф компаниям\firefox\geckodriver.exe",
        options=options
    )

    driver.get('https://www.rusprofile.ru/search-advanced')
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

    return driver

def sort_and_find(driver, region_input_def, activity_input_def, legal_form_input_def):
    try:
        if check_exists_by_xpath('/html/body/div[6]/div/div/div/div/div[6]', driver) == True:
            btn_activate = driver.find_element(By.XPATH, '/html/body/div[6]/div/div/div/div/div[6]')    
            btn_activate.click()
            time.sleep(7)

        # Сортировка по регионам и областям
        region_list = driver.find_elements(By.TAG_NAME, 'fieldset')[4]
        region_list.click()
        time.sleep(2)

        driver.execute_script("""
            const xz = Array.from(document.querySelector('#additional-search-region').children )[6] 
            const regionLi = xz.querySelector('div')
            regionLi.classList.add('expanded') 
        """)
        time.sleep(2)

        region_input = driver.find_element(By.XPATH, region_input_def)
        region_excel = region_input.text
        region_input.click()
        time.sleep(2)

        btn_region_input = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[5]/div[2]/div[2]/div[2]/div[3]/div[2]')
        btn_region_input.click()
        time.sleep(2)

        # Сортировка по Виду деятельности

        activity_list = driver.find_elements(By.TAG_NAME, 'fieldset')[2]
        activity_list.click()
        time.sleep(2)

        activity_input = driver.find_element(By.XPATH, activity_input_def)
        activity_excel = activity_input.text
        activity_input.click()
        time.sleep(2)

        btn_activity_input = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[3]/div[2]/div[2]/div[2]/div[3]/div[3]')
        btn_activity_input.click()
        time.sleep(2)

        # Сортировка по Правовой форме

        legal_form_list= driver.find_elements(By.TAG_NAME, 'fieldset')[8]
        legal_form_list.click()
        time.sleep(2)

        legal_form_input = driver.find_element(By.XPATH, legal_form_input_def)
        legal_form_excel = legal_form_input.text
        legal_form_input.click()
        time.sleep(2)

        btn_legal_form_input = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[9]/div[2]/div[2]/div[2]/div[4]/div[2]')
        btn_legal_form_input.click()
        time.sleep(2)

        # Сортировка по номерам

        telephone_list= driver.find_elements(By.TAG_NAME, 'fieldset')[9]
        telephone_list.click()
        time.sleep(2)

        btn_telephone_input = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/aside/form/fieldset[10]/div[2]')
        btn_telephone_input.click()
        time.sleep(2)

        sort_revenue = driver.find_element(By.XPATH, '//*[@id="search-sort-select"]')
        sort_revenue.click()
        time.sleep(2)
        sort_revenue.click()
        time.sleep(2)

        # Навигация по страницам
        list_links = []


        for i in range(1, 150000):
            print(i)
            btn_list_ul = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/section/div/div[5]/ul')
            btn_list_li = btn_list_ul.find_elements(By.TAG_NAME, 'li')
            if driver.find_element(By.XPATH, '//*[@id="pager-holder"]').get_attribute('class') == 'pager-holder hidden':
                find_link(list_links, driver)
                break
            # print(len(btn_list_li))
            if len(btn_list_li) == 4:
                btn_list_4 = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/section/div/div[5]/ul/li[4]/span')
                # print(btn_list_5)
                if btn_list_4.get_attribute('class') == 'fakelink nav-arrow nav-next disabled':
                    break
                btn_list_4.click()
                time.sleep(3)
                find_link(list_links, driver)
                # print('1')
            elif len(btn_list_li) == 5:
                btn_list_5 = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/section/div/div[5]/ul/li[5]/span')
                # print(btn_list_5)
                if btn_list_5.get_attribute('class') == 'fakelink nav-arrow nav-next disabled':
                    break
                btn_list_5.click()
                time.sleep(3)
                find_link(list_links, driver)
                # print('1')
            elif len(btn_list_li) == 6:
                btn_list_6 = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/section/div/div[5]/ul/li[6]/span')
                btn_list_6.click()
                time.sleep(3)
                company_items = driver.find_element(By.CLASS_NAME, 'additional-results')
                company_item_list = company_items.find_elements(By.TAG_NAME, 'a')
                for i in company_item_list:
                    list_links.append(i.get_attribute('href'))
                # print('2')
            elif len(btn_list_li) == 7:
                btn_list_7 = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div/section/div/div[5]/ul/li[7]/span')
                btn_list_7.click()
                time.sleep(3)
                find_link(list_links, driver)
                    # print('3')
        time.sleep(3)

        print(list_links)

        create_Excel(list_links, region_excel, activity_excel, legal_form_excel)
    except:
        pass
    finally:
        driver.close()
        driver.quit()






sort_and_find(authentication(),
    region_input_def,
    activity_input_def,
    legal_form_one_input_def
)
time.sleep(2)
sort_and_find(authentication(), 
    region_input_def,
    activity_input_def,
    legal_form_two_input_def
)
time.sleep(2)
sort_and_find(authentication(), 
    region_input_def,
    activity_input_def,
    legal_form_three_input_def
)
time.sleep(2)
sort_and_find(authentication(), 
    region_input_def,
    activity_input_def,
    legal_form_four_input_def
)