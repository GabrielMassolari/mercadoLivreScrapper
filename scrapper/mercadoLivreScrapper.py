from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import pandas as pd
import os
import time


class MercadoLivreScrapper:
    def __init__(self, filename):
        service = Service(executable_path='C:\Programação\Python\mercadoLivreScrapper\chromedriver.exe')
        self.__driver = webdriver.Chrome(service=service)
        self.__filename = ''
        self.set_filename(filename)

    def get_filename(self):
        return self.__filename

    def set_filename(self, value):
        if not os.path.isfile(value) or not value.split('.')[-1] in ['xlsx', 'xls']:
            raise FileNotFoundError(f"{value} file was not found or is not a excel file")
        self.__filename = value

    def read_products_from_xlsx(self):
        print(self.__filename)
        dataframe = pd.read_excel(f"C:\Programação\Python\mercadoLivreScrapper\{self.__filename}")

        if not dataframe.columns == ['Product']:
            raise Exception(f"{self.__filename} is not in pattern (Need has only 'Product' column)")

        return dataframe['Product'].to_list()

    def search_items_from_products_list(self):
        self.__driver.get("https://www.mercadolivre.com.br/")
        results = []
        products = self.read_products_from_xlsx()

        for product in products:
            nav_search = self.__driver.find_element(By.NAME, "as_word")
            nav_search.clear()
            nav_search.send_keys(product)
            nav_search.send_keys(Keys.RETURN)

            time.sleep(2)

            price = self.__driver.find_element(By.CLASS_NAME, 'andes-money-amount__fraction').text
            print(price)

            results.append({'Product': product, 'Price': price})

        print(results)

        self.__driver.close()
