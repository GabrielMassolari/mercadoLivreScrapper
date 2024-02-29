from scrapper.mercadoLivreScrapper import MercadoLivreScrapper
if __name__ == "__main__":
    scraper = MercadoLivreScrapper("ProdutosMercadoLivre.xlsx")
    scraper.search_items_from_products_list()
