from selenium import webdriver
from selenium.webdriver.chrome.service import Service

from bs4 import BeautifulSoup
import time
from openpyxl import Workbook
from urllib.parse import urljoin
import os

options = webdriver.ChromeOptions()
options.add_argument('--headless')
chrome_prefs = {}
options.experimental_options["prefs"] = chrome_prefs
chrome_prefs["profile.default_content_settings"] = {"images": 2}
chrome_prefs["profile.managed_default_content_settings"] = {"images": 2}

ser = Service("chromedriver.exe")
browser = webdriver.Chrome(service=ser, options=options)


def scrolling_page():
    browser.execute_script("window.scrollTo(0, 50);")
    browser.implicitly_wait(1)
    time.sleep(0.5)
    browser.execute_script("window.scrollTo(0, 500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 1000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 1500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 2000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 2500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 3000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 3500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 4000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 4500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 5000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 5500);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 6000);")
    browser.implicitly_wait(1)
    time.sleep(0.2)
    browser.execute_script("window.scrollTo(0, 6500);")
    browser.implicitly_wait(1)
    time.sleep(0.5)


wb = Workbook()
sheet = wb.active
sheet.title = "Silpo"
sheet['A1'] = 'Назва'
sheet['B1'] = 'Ціна'
sheet['C1'] = 'Ціна без знижки'
row = 2

urls = ['https://shop.silpo.ua/category/22',
        'https://shop.silpo.ua/category/476',
        'https://shop.silpo.ua/category/433',
        'https://shop.silpo.ua/category/277',
        'https://shop.silpo.ua/category/374',
        'https://shop.silpo.ua/category/316',
        'https://shop.silpo.ua/category/1468',
        'https://shop.silpo.ua/category/486',
        'https://shop.silpo.ua/category/359',
        'https://shop.silpo.ua/category/234',
        'https://shop.silpo.ua/category/264',
        'https://shop.silpo.ua/category/65',
        'https://shop.silpo.ua/category/130',
        'https://shop.silpo.ua/category/498',
        'https://shop.silpo.ua/category/308',
        'https://shop.silpo.ua/category/52',
        'https://shop.silpo.ua/category/470',
        'https://shop.silpo.ua/category/535',
        'https://shop.silpo.ua/category/567',
        'https://shop.silpo.ua/category/449',
        'https://shop.silpo.ua/category/653',
        'https://shop.silpo.ua/category/1477']
for url in urls:
    browser.get(url)
    print(url)
    time.sleep(0.5)
    scrolling_page()
    generated_html = browser.page_source
    soup = BeautifulSoup(generated_html, 'lxml')
    browser.implicitly_wait(1)
    time.sleep(0.5)
    if soup.find_all("div", class_='pagination-link'):
        last_div = soup.find_all("div", class_='pagination-link', text=True)[-1]
        page = last_div.find(text=True)
        print(last_div)
        print(page)
        count = 1
        while count <= int(page):
            base_url = urljoin(url, '?to=%s&from=%s')
            base_url = base_url % (int(count), int(count))
            browser.get(base_url)
            scrolling_page()
            browser.implicitly_wait(2)
            time.sleep(0.5)
            generated_html = browser.page_source

            soup = BeautifulSoup(generated_html, 'lxml')
            product_name = soup.find_all('div', class_='lazyload-wrapper')
            for n, i in enumerate(product_name, start=1):
                if i.find('div', class_='product-title') is None:
                    itemName = None
                else:
                    itemName = i.find('div', class_='product-title').text
                if i.find('div', class_='current-integer') is None:
                    itemFirstPrice = None
                else:
                    itemFirstPrice = i.find('div', class_='current-integer').text
                if i.find('div', class_='current-fraction') is None:
                    itemSecondPrice = None
                else:
                    itemSecondPrice = i.find('div', class_='current-fraction').text
                if i.find('div', class_='old-integer') is None:
                    itemPriceOldFirst = None
                else:
                    itemPriceOldFirst = i.find('div', class_='old-integer').text
                if i.find('div', class_='old-fraction') is None:
                    itemPriceOldSecond = None
                else:
                    itemPriceOldSecond = i.find('div', class_='old-fraction').text
                Price = ".".join((str(itemFirstPrice), str(itemSecondPrice)))
                OldPrice = ".".join((str(itemPriceOldFirst), str(itemPriceOldSecond)))
                sheet['A' + str(row)] = itemName
                sheet['B' + str(row)] = Price
                sheet['C' + str(row)] = OldPrice
                row += 1
                print(f'{n}:  {itemName} - {str(itemFirstPrice)}.{str(itemSecondPrice)} - {str(itemPriceOldFirst)}.{str(itemPriceOldSecond)}')
            print(base_url)
            print(count)
            count += 1
    else:
        soup = BeautifulSoup(generated_html, 'lxml')
        product_name = soup.find_all('div', class_='lazyload-wrapper')
        for n, i in enumerate(product_name, start=1):
            if i.find('div', class_='product-title') is None:
                itemName = None
            else:
                itemName = i.find('div', class_='product-title').text
            if i.find('div', class_='current-integer') is None:
                itemFirstPrice = None
            else:
                itemFirstPrice = i.find('div', class_='current-integer').text
            if i.find('div', class_='current-fraction') is None:
                itemSecondPrice = None
            else:
                itemSecondPrice = i.find('div', class_='current-fraction').text
            if i.find('div', class_='old-integer') is None:
                itemPriceOldFirst = None
            else:
                itemPriceOldFirst = i.find('div', class_='old-integer').text
            if i.find('div', class_='old-fraction') is None:
                itemPriceOldSecond = None
            else:
                itemPriceOldSecond = i.find('div', class_='old-fraction').text
            Price = ".".join((str(itemFirstPrice), str(itemSecondPrice)))
            OldPrice = ".".join((str(itemPriceOldFirst), str(itemPriceOldSecond)))
            sheet['A' + str(row)] = itemName
            sheet['B' + str(row)] = Price
            sheet['C' + str(row)] = OldPrice
            row += 1
            print(f'{n}:  {itemName} - {str(itemFirstPrice)}.{str(itemSecondPrice)} - {str(itemPriceOldFirst)}.{str(itemPriceOldSecond)}')



filename = 'Siplo_shop.xlsx'
wb.save(filename=filename)
os.chdir(os.getcwd())
###Windows###
os.system('start excel.exe "%s\\%s"' %(os.getcwd(), filename,))
##Ubuntu###
#os.system('/usr/bin/libreoffice --calc "%s\\%s"' %(os.getcwd(), filename, ))
browser.quit()