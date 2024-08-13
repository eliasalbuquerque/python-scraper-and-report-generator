# title: 'app'
# author: 'Elias Albuquerque'
# version: '0.3.0'
# created: '2024-08-08'
# update: '2024-08-10'


import logging
import json
import datetime
import os
import tempfile
import shutil
from time import sleep
from src.settings import Settings
from src.website import Website
from src.office import Office
from src.report import report_content


def get_current_date_time():
    now = datetime.datetime.now()
    return now


def load_config():
    with open('config.json', 'r', encoding='utf8') as file:
        return json.load(file)


def string_to_float_to_string(value):
    """
    Converte uma string para um valor float com duas casas decimais.
    """

    value = value.replace(",", ".")
    float_value = float(value)
    float_value_rounded = round(float_value, 2)
    value_to_string = str(float_value_rounded)
    value_rounded = value_to_string.replace(".", ",")

    return value_rounded


def main():
    # Criar a pasta temporária para salvar o screenshot
    tempdir = tempfile.mkdtemp()

    try:
        # 0. Carrega as configuracoes e variaveis da aplicacao
        settings = Settings()
        config = load_config()
        url = config['website']['url']
        xp_button_cookie = config['website']['xp_button_cookie']
        xp_quote = config['website']['xp_quote']
        now = get_current_date_time()
        today = now.strftime("%d/%m/%Y")
        hour = now.strftime("%H:%M:%S")
        author = config['office']['author']
        if author == "null":
            author = os.getlogin().capitalize()
        report_file = "relatorio-" + now.strftime("%Y%m%d-%H%M%S")
        report_path = os.path.join('reports', report_file)


        # 1. Acessar o site e extrair o valor da cotacao
        website = Website(settings)
        website.access_website(url)
        website.click_on_element(xp_button_cookie)
        website.zoom_out_of_website(86)
    
        quote = website.extract_text_from_element(xp_quote, 'cotação')
        screenshot = website.take_screenshot(tempdir)
    

        # 2. Criar arquivo Word.docx montar o relatorio
        office = Office()
        office.create_document(report_file, report_path)


        # 3. Adicionar conteudo no relatorio
        quote = "R$ " + string_to_float_to_string(quote)
        report_content(office, report_path, quote, today, hour, url, screenshot, author)


        # 4. Transforme em um PDF
        office.convert_docx_to_pdf(report_path)


        # * revise as docstrings de todos os modulos
        # * informar que o modulo report_content é necessario refatorar o codigo caso seja reutilizado em outro projeto


        # 5. Entrega como executável
        # - Transforme o código em um instalador(instruções nas dicas abaixo)

    finally:
        # Remover a pasta temporária no final da execução, se necessário
        shutil.rmtree(tempdir)

if __name__ == '__main__':
    main()
