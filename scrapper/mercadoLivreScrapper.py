from selenium import webdriver
import pandas as pd
import os


class MercadoLivreScrapper:
    def __init__(self, filename):
        self.driver = webdriver.Chrome()
        self.filename = filename

    @property
    def filename(self):
        return self.filename

    @filename.setter
    def filename(self, value):
        if os.path.isfile(value):
            raise FileNotFoundError(f"${value} file was not found")
        self.filename = value

    def read_items_from_xlsx(self):
        dataframe = pd.read_excel(self.filename, encoding='utf-8')

        if not dataframe.columns == ['Product']:
            raise Exception(f"${self.filename} is not in pattern (Need has only 'Product' column)")

        return dataframe['Product'].to_list()

