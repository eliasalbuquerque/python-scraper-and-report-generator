# title: 'module settings to inittiate logging and webdriver'
# author: 'Elias Albuquerque'
# version: '0.1.0'
# created: '2024-08-08'
# update: '2024-08-08'

"""
Este módulo configura as configurações básicas da aplicação, incluindo:

- Logging: Configura o sistema de logging usando o arquivo 'config.ini'.
- Driver do Chrome: Configura o driver do Chrome com as opções desejadas.
- Wait: Configura o objeto WebDriverWait para lidar com a espera de elementos na página.

A classe Settings fornece métodos para acessar o driver do Chrome, o objeto wait e realizar o setup do logging.

Uso:

    from settings import Settings

    settings = Settings()

    # Acessando o driver do Chrome
    driver = settings.driver

    # Acessando o objeto wait
    wait = settings.wait

    # Usando o logging
    logging.info("Mensagem de log")
"""

import os
import logging.config
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException, ElementNotSelectableException


class Settings:
    def __init__(self):
        self._setup_logging()

        logging.info('Aplicação Iniciada.')
        logging.info('Iniciando configurações da aplicação...')
        
        self.driver = self._setup_driver()
        self.wait = self._setup_wait()

    def _setup_logging(self):
        """
        Configura o logging da aplicação utilizando o arquivo 'config.ini'.
        """
        
        current_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(current_dir, 'config.ini')
        logging.config.fileConfig(config_path, disable_existing_loggers=False)

    def _setup_driver(self):
        """Configura o driver do Chrome."""

        try:
            options = Options()
            arguments = [
                '--disable-notifications',
                '--block-new-web-contents',
                '--no-default-browser-check',
                '--lang=pt-BR',
                # '--headless',
                '--window-position=36,68',
                '--window-size=1100,750',]

            for argument in arguments:
                options.add_argument(argument)

            options.add_experimental_option("excludeSwitches", ["enable-logging"])

            driver = webdriver.Chrome(options=options)
            return driver

        except FileNotFoundError:
            logging.error('Driver do Chrome não encontrado. Verifique a instalação do driver ou baixe em: https://developer.chrome.com/docs/chromedriver/downloads')
            return None

        except Exception as e:
            logging.error(f'Erro na configuração do driver: {e}')
            return None

    def _setup_wait(self):
        """Configura o wait."""

        wait = WebDriverWait(
            self.driver,
            15,
            poll_frequency=1,
            ignored_exceptions=[
                NoSuchElementException,
                ElementNotVisibleException,
                ElementNotSelectableException])
        return wait
