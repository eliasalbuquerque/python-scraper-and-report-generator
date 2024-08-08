# title: 'website'
# author: 'Elias Albuquerque'
# version: '0.1.0'
# created: '2024-08-08'
# update: '2024-08-08'

"""
"""

import logging
import tempfile
from time import sleep
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Website:
    def __init__(self, settings):
        self.driver = settings.driver
        self.wait = settings.wait

    def access_website(self, url):
        """
        Esta função acessa um site especificado usando um driver de navegador.
        Args: url (str), driver (webdriver), wait (int)
        Returns: driver (webdriver)
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
        Esta função clica em um elemento específico na página.
        Args: xpath_element (str)
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
        Esta função extrai dados de um elemento específico na página.
        Args: xpath_element (str)
        Returns: str
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
        Esta função extrai dados de um elemento específico na página.
        Args: xpath_element (str)
        Returns: str
        """

        logging.info(f'Extraindo {data_to_extract} do site...')

        try:
            element = self.wait.until(EC.visibility_of_element_located((
                By.XPATH, xp_current_temperature))).get_attribute(attribute)
            return element.text

        except TimeoutException as e:
            logging.error(f'Erro ao extrair {data_to_extract} do elemento: {e}')
            return None


    def zoom_out_of_website(self, zoom_out_percentage):
        """
        Esta função aplica um zoom out na página.
        Args: zoom_out_percentage (float)
        Ex. website.zoom_out_of_website(67) # Aplica um zoom out de 67% na pagina
        """

        try:
            zoom_value = (zoom_out_percentage / 100)
            self.driver.execute_script(f'document.body.style.zoom="{zoom_value}"')
            sleep(5)

        except WebDriverException as e:
            logging.error(f'Erro ao aplicar zoom out: {e}')


    def take_screenshot(self):
        """
        Esta função captura uma screenshot da página atual.
        Args: screenshot_path (str)
        Returns: None
        """

        try:
            screenshot = self.driver.get_screenshot_as_png()
            return screenshot

        except WebDriverException as e:
            logging.error(f'Erro ao capturar a screenshot: {e}')