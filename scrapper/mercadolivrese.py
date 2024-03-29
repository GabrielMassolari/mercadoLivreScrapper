from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import utils
import smtplib
import pandas as pd
import os


class MercadoLivreSe:
    """Class that search mercado livre product informations using selenium lib"""
    def __init__(self, filename):
        service = Service(executable_path=utils.get_complete_root_file_path('chromedriver.exe'))
        self.__driver = webdriver.Chrome(service=service)
        self.__filename = ''
        self.set_filename(filename)

    def get_filename(self):
        """Return object filename attr"""
        return self.__filename

    def set_filename(self, value):
        """Set object filename attr with excel extension validation"""
        if not os.path.isfile(value) or not value.split('.')[-1] in ['xlsx', 'xls']:
            raise FileNotFoundError(f"{value} file was not found or is not a excel file")
        self.__filename = value

    def read_products_from_excel_file(self):
        """Return de product name in excel file"""
        dataframe = pd.read_excel(utils.get_complete_root_file_path(self.__filename))

        if not list(dataframe.columns) == ['Product']:
            raise Exception(f"{self.__filename} is not in pattern (Need has only 'Product' column)")

        return dataframe['Product'].to_list()

    def search_products_info_in_mercado_livre(self, average_price_flag=False):
        """Return products info read by excel in mercado livre website using selenium"""
        self.__driver.get("https://www.mercadolivre.com.br/")
        results = []
        products = self.read_products_from_excel_file()

        for product in products:
            nav_search = self.__driver.find_element(By.NAME, "as_word")
            nav_search.clear()
            nav_search.send_keys(product)
            nav_search.send_keys(Keys.RETURN)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.NAME, "as_word"))
            )

            try:
                element = self.__driver.find_element(By.CSS_SELECTOR, 'span.ui-search-price__part--medium .andes-money-amount__fraction')
            except NoSuchElementException:
                results.append({'Product': f'{product} | PRODUCT NOT FOUND', 'Price': None})
                continue

            if average_price_flag:
                prices = self.__driver.find_elements(By.CSS_SELECTOR,
                                                     'span.ui-search-price__part--medium .andes-money-amount__fraction')
                prices = [int(element.text.replace(".", "")) for element in prices[:10]]
                average_price = sum(prices) / len(prices)

                results.append({'Product': product, 'AvgPrice': average_price})
            else:
                price = int(self.__driver.find_element(By.CSS_SELECTOR, 'span.ui-search-price__part--medium .andes-money-amount__fraction')
                            .text.replace(".", ""))

                results.append({'Product': product, 'Price': price})

        self.__driver.close()

        return results

    def save_products_info_in_excel(self, products_info):
        """Save products info in the same excel file (Overwrithe the content)"""
        dataframe = pd.DataFrame(products_info)
        dataframe.loc['Total'] = dataframe.sum(numeric_only=True)

        dataframe.to_excel(self.__filename, index=False)

    def send_email_with_excel_attachment(self):
        """Send email with excel attachment with mercado livre portuguese informations like subject and body"""
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

    def fill_spreadsheet_products_info_and_send_it_by_email(self, average_price_flag=False):
        """Function that realize all scrapping function like get products name,
        save new products info in excel and sent it by email"""
        products_info = self.search_products_info_in_mercado_livre(average_price_flag)
        self.save_products_info_in_excel(products_info)
        self.send_email_with_excel_attachment()
