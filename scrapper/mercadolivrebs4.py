from bs4 import BeautifulSoup
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import requests
import utils
import smtplib
import pandas as pd
import os


class MercadoLivreBS4:
    """Class that search mercado livre product informations using beautifulsoup4 lib"""
    def __init__(self, filename):
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
        """Return products info read by excel in mercado livre website using requests and bs4 lib"""
        results = []
        products = self.read_products_from_excel_file()
        url = "https://lista.mercadolivre.com.br/"

        for product in products:
            product_url_name = product.replace(" ", "-")
            complete_url = url + product_url_name

            params = {
                "sec-ch-ua": "'Google Chrome';v='87', 'Not;A Brand';v='99', 'Chromium';v='87'",
                "sec-ch-ua-mobile": "?0",
                "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36"
            }

            page = requests.get(complete_url, params)
            soup = BeautifulSoup(page.text, 'html.parser')

            price_element = soup.select_one('span.ui-search-price__part--medium .andes-money-amount__fraction')

            if not price_element:
                results.append({'Product': f'{product} | PRODUTO NAO ENCONTRADO', 'Price': None})
                continue

            if average_price_flag:
                prices = soup.select('span.ui-search-price__part--medium .andes-money-amount__fraction')
                prices = [int(element.text.replace(".", "")) for element in prices[:10]]
                average_price = sum(prices) / len(prices)

                results.append({'Product': product, 'AvgPrice': average_price})
            else:
                price = int(soup.select_one('span.ui-search-price__part--medium .andes-money-amount__fraction').text.replace(".", ""))

                results.append({'Product': product, 'Price': price})

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
