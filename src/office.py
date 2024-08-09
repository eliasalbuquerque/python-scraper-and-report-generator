# title: 'module office'
# author: 'Elias Albuquerque'
# version: '0.1.0'
# created: '2024-08-09'
# update: '2024-08-09'

import logging
import docx
import re
import os
from docx import Document


class Office:
    """
    Classe para gerenciar documentos do Office, com métodos para criar e manipular documentos.
    """

    def __init__(self):
        self.doc = None

    def create_document(self, file_name, file_path):
        """Cria um novo documento Word."""

        logging.info(f'Criando arquivo "{file_name}" ...')

        # Remove caracteres inválidos do nome do arquivo
        file_name = re.sub(r'[^a-zA-Z0-9_\.\-]', '', file_name)
        if not os.path.isdir(file_path):
            raise ValueError("O caminho do documento é inválido.")

        try:
            self.doc = Document()
            file_path = os.path.join(file_path, file_name)
            self.doc.save(file_path)

            return True

        except Exception as e:
            logging.error(f'Erro ao criar "{file_name}": {e}')
            return None
