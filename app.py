# title: 'app'
# author: 'Elias Albuquerque'
# version: '0.1.0'
# created: '2024-08-08'
# update: '2024-08-08'


import logging
import json
import datetime
from time import sleep
from src.settings import Settings
from src.website import Website
# from src.office import Office


def keep_the_window_open():  # teste
    logging.info('Mantendo site aberto por tempo indeterminado...')
    input('')


def get_current_date_time():
    now = datetime.datetime.now()
    return now.strftime("%Y-%m-%d %H:%M:%S")


def load_config():
    with open('config.json', encoding='utf8') as file:
        return json.load(file)


def main():
    settings = Settings()
    website = Website(settings)
    config = load_config()

    url = config['website']['url']
    xp_button_cookie = config['website']['xp_button_cookie']
    xp_quote = config['website']['xp_quote']

    website.access_website(url)
    website.click_on_element(xp_button_cookie)
    website.zoom_out_of_website(67)

    quote = website.extract_text_from_element(xp_quote, 'cotação')
    today = get_current_date_time()
    screenshot = website.take_screenshot()

    data_to_word_file = {
        "quote": quote,
        "today": today,
        "url": url,
        "screenshot": screenshot
    }

    # 2. Criação de arquivo word
    # - Organizar e armazenar os dados coletados de forma estruturada em um arquivo word
    # - Seu relatório(arquivo word) deve seguir este modelo de arquivo word conter todas essas informações aqui


    # 3. Transforme em um PDF
    # - Encontre uma forma de transformar o arquivo word em um arquivo python, usando somente python e transforme o arquivo word em um arquivo pdf
    # 4. Entrega como executável
    # - Transforme o código em um instalador(instruções nas dicas abaixo)

    keep_the_window_open()

if __name__ == '__main__':
    main()
