# title: 'module office'
# author: 'Elias Albuquerque'
# version: '0.3.0'
# created: '2024-08-09'
# update: '2024-08-13'


import logging
import docx
import re
import os
import shutil
import subprocess
import urllib.request
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class Office:
    """
    Classe para gerenciar a conversão de documentos .docx para .pdf usando
    diferentes ferramentas (Word, LibreOffice, Pandoc).

    Métodos:
    - create_document(file_name, file_path): Cria um novo arquivo.
    - convert_docx_to_pdf(docx_path): Converte um arquivo .docx para .pdf.
    - add_hyperlink(paragraph, url, text, color="0000FF", underline=True): Adicionar hiperlink em um elemento contido em um parágrafo
    """

    def __init__(self):
        self.doc = None

    def create_document(self, file_name, file_path):
        """Cria um novo documento com um nome e caminho definido.

        Args:
            file_name (str): O nome do arquivo a ser criado.
            file_path (str): O caminho completo para o arquivo, incluindo o nome do arquivo.
        """

        logging.info(f'Criando arquivo "{file_name}" ...')

        # Verifica se o diretório existe
        directory = os.path.dirname(file_path)
        if not os.path.exists(directory):
            try:
                os.makedirs(directory)
                logging.info(f'Diretório "{directory}" criado.')
            except Exception as e:
                logging.error(f'Erro ao criar diretório "{directory}": {e}')
                return None
            
        try:
            Document().save(file_path)

        except Exception as e:
            logging.error(f'Erro ao criar "{file_name}": {e}')
            return None

    def convert_docx_to_pdf(self, docx_path):
        """
        Converte um arquivo .docx para .pdf utilizando Word, LibreOffice ou Pandoc,
        conforme a disponibilidade das ferramentas.

        Parâmetros:
        - docx_path (str): Caminho completo para o arquivo .docx que deve ser convertido.

        Retorna:
        - bool: True se a conversão foi bem-sucedida, False se falhar em todas as tentativas.
        """

        program_files_path = os.environ.get('ProgramFiles', 'C:\\Program Files')

        try:
            exec_word = self._check_word_installed(program_files_path)
            if exec_word:
                self._convert_with_word(docx_path, exec_word)
                return True
        except PermissionError:
            logging.error("Permissão negada para acessar o Microsoft Word.")
        except Exception as e:
            logging.error(f"Erro ao converter com Word: {e}")

        try:
            exec_libreoffice = self._check_libreoffice_installed(program_files_path)
            if exec_libreoffice:
                self._convert_with_libreoffice(docx_path, exec_libreoffice)
                return True
        except Exception as e:
            logging.error(f"Erro ao converter com LibreOffice: {e}")
        
        try:
            exec_pandoc = self._check_pandoc_installed()
            if exec_pandoc:
                self._convert_with_pandoc(docx_path, exec_pandoc)
                return True
        except Exception as e:
            logging.error(f"Erro ao converter com Pandoc: {e}")

        return False

    def _check_word_installed(self, program_files_path):
        """
        Verifica se o Microsoft Word está instalado e retorna o caminho do executável.

        Parâmetros:
        - program_files_path (str): Caminho para o diretório "Arquivos de Programas".

        Retorna:
        - str: Caminho completo para o executável do Word se encontrado, None caso contrário.
        """
        logging.debug('Verificando se o Microsoft Word está instalado...')
        path = os.path.join(program_files_path, 'Microsoft Office', 'root', 'Office16')

        for root, _, files in os.walk(path):
            if 'WINWORD.EXE' in files:
                return os.path.join(root, 'WINWORD.EXE')

        try:
            output = subprocess.check_output(['where', 'WINWORD.EXE'], shell=True, text=True)
            logging.debug(f'Microsoft Word instalado: {output}')
            return output.strip()
        except subprocess.CalledProcessError:
            logging.debug('Microsoft Word não está instalado')
            return None

    def _check_libreoffice_installed(self, program_files_path):
        """
        Verifica se o LibreOffice está instalado e retorna o caminho do executável.

        Parâmetros:
        - program_files_path (str): Caminho para o diretório "Arquivos de Programas".

        Retorna:
        - str: Caminho completo para o executável do LibreOffice se encontrado, None caso contrário.
        """
        logging.debug('Verificando se o LibreOffice está instalado...')
        path = os.path.join(program_files_path, 'LibreOffice')

        for root, _, files in os.walk(path):
            if 'soffice.exe' in files:
                return os.path.join(root, 'soffice.exe')

        return None

    def _check_pandoc_installed(self):
        """
        Verifica se o Pandoc está instalado e retorna o caminho do executável.

        Retorna:
        - str: Caminho completo para o executável do Pandoc se encontrado, None caso contrário.
        """
        logging.debug('Verificando se o Pandoc está instalado...')

        # Verifica no PATH do sistema
        try:
            output = subprocess.check_output(['where', 'pandoc'], shell=True, text=True).strip()
            if output:
                logging.debug(f'Pandoc encontrado no PATH: {output}')
                return output
        except subprocess.CalledProcessError:
            logging.debug('Pandoc não encontrado no PATH.')

        # Verifica na pasta AppData\Local\Pandoc
        appdata_path = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Pandoc', 'pandoc.exe')
        if os.path.exists(appdata_path):
            logging.debug(f'Pandoc encontrado em AppData: {appdata_path}')
            return appdata_path

        # Verifica na pasta Program Files
        program_files_path = os.environ.get('ProgramFiles', 'C:\\Program Files')
        program_files_pandoc = os.path.join(program_files_path, 'Pandoc', 'pandoc.exe')
        if os.path.exists(program_files_pandoc):
            logging.debug(f'Pandoc encontrado em Program Files: {program_files_pandoc}')
            return program_files_pandoc

        # Se não encontrado, tenta instalar
        logging.debug('Pandoc não está instalado.')
        pandoc = PandocInstaller()
        if pandoc.install():
            # Tenta novamente encontrar após a instalação
            return self._check_pandoc_installed()

        return None

    def _convert_with_word(self, docx_path, exec_word):
        """
        Converte um arquivo .docx para .pdf usando o Microsoft Word.

        Parâmetros:
        - docx_path (str): Caminho completo para o arquivo .docx.
        - exec_word (str): Caminho completo para o executável do Microsoft Word.

        Retorna:
        - None
        """
        pdf_path = docx_path.replace('.docx', '.pdf')
        subprocess.run([exec_word, docx_path, '/t', pdf_path, '/q'])
        logging.info(f'Arquivo convertido para PDF: {pdf_path}')

    def _convert_with_libreoffice(self, docx_path, exec_libreoffice):
        """
        Converte um arquivo .docx para .pdf usando o LibreOffice.

        Parâmetros:
        - docx_path (str): Caminho completo para o arquivo .docx.
        - exec_libreoffice (str): Caminho completo para o executável do LibreOffice.

        Retorna:
        - None
        """
        pdf_path = docx_path.replace('.docx', '.pdf')
        subprocess.run([exec_libreoffice, '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(docx_path), docx_path])
        logging.info(f'Arquivo convertido para PDF: {pdf_path}')

    def _convert_with_pandoc(self, docx_path, exec_pandoc):
        """
        Converte um arquivo .docx para .pdf usando o Pandoc.

        Parâmetros:
        - docx_file (str): Caminho completo para o arquivo .docx.

        Retorna:
        - None
        """
        pdf_path = docx_path.replace('.docx', '.pdf')
        subprocess.run([exec_pandoc, docx_path, '-o', pdf_path])
        logging.info(f'Arquivo convertido para PDF: {pdf_path}')


    def add_hyperlink(self, paragraph, url, text, color="0000FF", underline=True):
        """
        Adiciona um hyperlink a um parágrafo.

        :param paragraph: O parágrafo onde o hyperlink será adicionado.
        :param url: O URL para o qual o hyperlink deve apontar.
        :param text: O texto do hyperlink.
        :param color: A cor do texto do hyperlink como string hexadecimal (sem o '#').
        :param underline: Se o hyperlink deve ser sublinhado ou não.
        """

        # Garante que o URL comece com http:// ou https://
        if not url.startswith("http://") and not url.startswith("https://"):
            url = "http://" + url

        # Cria o elemento de relacionamento do hyperlink
        part = paragraph.part
        r_id = part.relate_to(url, qn('r:hyperlink'), is_external=True)

        # Cria o elemento hyperlink
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)

        # Cria um novo elemento run para o texto do hyperlink
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')

        # Adiciona cor e sublinhado ao hyperlink
        color_element = OxmlElement('w:color')
        color_element.set(qn('w:val'), color)
        rPr.append(color_element)

        if underline:
            underline_element = OxmlElement('w:u')
            underline_element.set(qn('w:val'), 'single')
            rPr.append(underline_element)

        new_run.append(rPr)

        # Adiciona o texto do hyperlink
        text_element = OxmlElement('w:t')
        text_element.text = text
        new_run.append(text_element)

        # Adiciona o run ao hyperlink
        hyperlink.append(new_run)

        # Adiciona o hyperlink ao parágrafo
        paragraph._element.append(hyperlink)

        return paragraph


class PandocInstaller:
    """
    Classe responsável pela instalação do Pandoc, verificando e instalando
    o winget se necessário.

    Métodos:
    - install(): Verifica e instala o winget, depois instala o Pandoc.
    """

    def install(self):
        """
        Verifica se o winget e o Pandoc estão instalados. Caso o winget não esteja
        presente, ele é instalado. Em seguida, o Pandoc é instalado via winget.

        Retorna:
        - bool: True se a instalação do Pandoc foi bem-sucedida, False caso contrário.
        """
        try:
            if not self._is_winget_installed():
                logging.info("winget não encontrado. Instalando winget...")
                self._install_winget()
            
            if self._is_pandoc_installed():
                logging.info("Pandoc já está instalado.")
                return True

            logging.info("Instalando Pandoc via winget...")
            subprocess.check_call(["winget", "install", "--id", "JohnMacFarlane.Pandoc", "-e", "--source", "winget"])
            logging.info("Instalação do Pandoc concluída.")
            return True
        except Exception as e:
            logging.error(f"Erro durante a instalação: {e}")
            return False

    def _is_winget_installed(self):
        """
        Verifica se o winget está instalado na máquina.

        Retorna:
        - bool: True se o winget está instalado, False caso contrário.
        """
        try:
            subprocess.check_output(["winget", "--version"], text=True)
            return True
        except (FileNotFoundError, subprocess.CalledProcessError):
            return False

    def _install_winget(self):
        """
        Baixa e instala o winget, se não estiver presente na máquina.

        Exceções:
        - Exception: Levantada se houver falha ao baixar ou instalar o winget.
        """
        try:
            url = "https://aka.ms/getwinget"
            installer_path = "wingetInstaller.msixbundle"
            urllib.request.urlretrieve(url, installer_path)

            logging.info("Executando o instalador do winget...")
            subprocess.check_call(["powershell", "Add-AppxPackage", "-Path", installer_path])
            logging.info("winget instalado com sucesso.")
            os.remove(installer_path)
        except Exception as e:
            raise Exception(f"Falha ao instalar o winget: {e}")

    def _is_pandoc_installed(self):
        """
        Verifica se o Pandoc está instalado e na versão correta.

        Retorna:
        - bool: True se o Pandoc está instalado e na versão 2.19.1, False caso contrário.
        """
        try:
            output = subprocess.check_output(["pandoc", "--version"], text=True)
            return True
        except FileNotFoundError:
            return False
        return False