# title: 'website'
# author: 'Elias Albuquerque'
# version: '0.1.1'
# created: '2024-08-08'
# update: '2024-08-09'

import logging
import tempfile
import os
from time import sleep
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class Website:
    """
    Classe para interagir com um site web usando Selenium.
    """

    def __init__(self, settings):
        """
        Inicializa a classe Website com o driver e a configuração de espera.
        Args:
            settings: Um objeto que contém as configurações do driver e da espera.
        """

        self.driver = settings.driver
        self.wait = settings.wait

    def access_website(self, url):
        """
        Acessa o site especificado usando o driver de navegador.
        Args:
            url (str): URL do site a ser acessado.
        Returns:
            selenium.webdriver.WebDriver: O driver de navegador, ou None caso ocorra um erro.
        """

        logging.info('Acessando o site (aguarde!)...')

        try:
            self.driver.get(url)
            sleep(15)
            return self.driver

        except TimeoutException as e:
            logging.error(f'Erro ao acessar o site {url}: {e}')
            return None
        except WebDriverException as e:
            logging.error(f'Erro ao acessar o site {url}: {e}')
            return None

    def click_on_element(self, xpath_element):
        """
        Clica em um elemento específico na página.
        Args:
            xpath_element (str): O XPath do elemento a ser clicado.
        """

        try:
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, xpath_element))
            )
            element.click()

        except TimeoutException as e:
            logging.error(f'Erro ao clicar no elemento: {e}')

    def extract_text_from_element(self, xpath_element, data_to_extract='dados'):
        """
        Extrai o texto de um elemento específico na página.
        Args:
            xpath_element (str): O XPath do elemento a ser extraído.
            data_to_extract (str, optional): Uma descrição do tipo de dado 
                sendo extraído. Padrão 'dados'.
        Returns:
            str: O texto do elemento, ou None caso ocorra um erro.
        """

        logging.info(f'Extraindo {data_to_extract} do site...')

        try:
            element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, xpath_element))
            )
            return element.text

        except TimeoutException as e:
            logging.error(f'Erro ao extrair {data_to_extract} do elemento: {e}')
            return None

    def extract_text_from_attribute_of_element(self, xpath_element, attribute, data_to_extract='dados'):
        """
        Extrai o valor de um atributo específico de um elemento na página.
        Args:
            xpath_element (str): O XPath do elemento.
            attribute (str): O nome do atributo a ser extraído.
            data_to_extract (str, optional): Uma descrição do tipo de dado 
                sendo extraído. Padrão 'dados'.
        Returns:
            str: O valor do atributo, ou None caso ocorra um erro.
        """

        logging.info(f'Extraindo {data_to_extract} do site...')

        try:
            element = self.wait.until(
                EC.visibility_of_element_located((By.XPATH, xpath_element))
            ).get_attribute(attribute)
            return element

        except TimeoutException as e:
            logging.error(f'Erro ao extrair {data_to_extract} do elemento: {e}')
            return None

    def zoom_out_of_website(self, zoom_out_percentage):
        """
        Aplica um zoom out na página.
        Args:
            zoom_out_percentage (float): A porcentagem de zoom out a ser aplicada.
        """

        try:
            zoom_value = (zoom_out_percentage / 100)
            self.driver.execute_script(
                f'document.body.style.zoom="{zoom_value}"')
            sleep(5)

        except WebDriverException as e:
            logging.error(f'Erro ao aplicar zoom out: {e}')

    def take_screenshot(self, tempdir):
        """
        Captura uma screenshot da página atual e salva em uma pasta temporária.
        Returns:
            str: O caminho completo do arquivo da screenshot, ou None caso ocorra um erro.
        """

        try:
            file_name = "screenshot.png"
            file_path = os.path.join(tempdir, file_name)

            self.driver.save_screenshot(file_path)

            return file_path

        except WebDriverException as e:
            logging.error(f"Erro ao capturar a screenshot: {e}")
            return None