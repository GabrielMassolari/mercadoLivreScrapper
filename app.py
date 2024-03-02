from scrapper.mercadoLivreScrapper import MercadoLivreScrapper
if __name__ == "__main__":
    scraper = MercadoLivreScrapper("ProdutosMercadoLivre.xlsx")
    scraper.fill_spreadsheet_products_info_and_send_it_by_email(average_price_flag=False)
