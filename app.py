from scrapper.mercadolivrese import MercadoLivreSe
from scrapper.mercadolivrebs4 import MercadoLivreBS4

if __name__ == "__main__":
    scraper = MercadoLivreSe("ProdutosMercadoLivre2.xlsx")
    scraper.fill_spreadsheet_products_info_and_send_it_by_email(average_price_flag=False)

    #scraper = MercadoLivreBS4("ProdutosMercadoLivre2.xlsx")
    #scraper.fill_spreadsheet_products_info_and_send_it_by_email(average_price_flag=False)
