from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import utils
import smtplib
import pandas as pd
import os
import time


class MercadoLivreScrapper:
    def __init__(self, filename):
        service = Service(executable_path=utils.get_complete_root_file_path('chromedriver.exe'))
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
        dataframe = pd.read_excel(utils.get_complete_root_file_path(self.__filename))

        if not dataframe.columns == ['Product']:
            raise Exception(f"{self.__filename} is not in pattern (Need has only 'Product' column)")

        return dataframe['Product'].to_list()

    def search_info_from_xlsx_products(self):
        self.__driver.get("https://www.mercadolivre.com.br/")
        results = []
        products = self.read_products_from_xlsx()

        for product in products:
            nav_search = self.__driver.find_element(By.NAME, "as_word")
            nav_search.clear()
            nav_search.send_keys(product)
            nav_search.send_keys(Keys.RETURN)

            time.sleep(2)

            price = int(self.__driver.find_element(By.CLASS_NAME, 'andes-money-amount__fraction').text.replace(".", ""))

            results.append({'Product': product, 'Price': price})

        self.__driver.close()

        return results

    def save_products_info_in_xlsx(self, products_info):
        dataframe = pd.DataFrame(products_info)
        dataframe.loc['Total'] = dataframe.sum(numeric_only=True)
        print(dataframe)

        dataframe.to_excel(self.__filename, index=False)

    def send_email_with_excel_atachment(self):
        load_dotenv()

        sender_email = os.getenv("SENDER_EMAIL")
        sender_password = os.getenv("GMAIL_PASSWORD")
        recipient_email = os.getenv("SENDER_EMAIL")
        subject = "Tabela Mercado Livre"
        body = "Segue em anexo a planilha de preços dos produtos do mercado livre"

        with open(utils.get_complete_root_file_path(self.__filename), "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={self.__filename}",
        )

        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = recipient_email

        html_part = MIMEText(body)
        message.attach(html_part)
        message.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())

    def fill_spreadsheet_products_info_and_send_it_by_email(self):
        products_info = self.search_info_from_xlsx_products()
        self.save_products_info_in_xlsx(products_info)
        self.send_email_with_excel_atachment()
