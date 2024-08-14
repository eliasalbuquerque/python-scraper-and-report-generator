# title: 'setup cx_freeze'
# author: 'Elias Albuquerque'
# version: '0.1.0'
# created: '2024-08-14'
# update: '2024-08-14'


import sys
import os
from cx_Freeze import setup, Executable

# Dados para o executável
executables = [Executable(
    script="app.py",
    base="Console",
)]

# Defina os pacotes necessários para a aplicação
packages = [
    "selenium", 
    "urllib.request", 
    "os", 
    "tempfile", 
    "shutil", 
    "pathlib", 
    "json",
    "datetime",
    "time",
    "re",
    "docx",
    "subprocess"
]

# Inclua os arquivos e diretórios necessários
include_files = [
    ("src", "src"), 
    ("config.json", "config.json"),
    ("LICENSE", "LICENSE"),
    ("README.md", "README.md")
]

# Inclua os arquivos do pacote "requirements.txt"
requirements = ["python-docx==1.1.2", "selenium==4.23.1", "cx_Freeze==7.2.0"]

# Crie o arquivo de configuração para o cx_Freeze
build_exe_options = {
    "packages": packages,
    "include_files": include_files,
    "include_msvcr": True,
    "silent_level": 3
}

# Crie o arquivo de configuração para o cx_Freeze
setup(
    name="PythonReport",
    version="1.0",
    author="Elias Albuquerque",
    description="Aplicação que automatiza a coleta de dados de cotação do dólar de um site, gera um relatório em PDF com data, cotação, captura de tela e autor do relatório.",
    url="https://github.com/eliasalbuquerque/python-scraper-and-report-generator",
    license="MIT",
    license_files=["LICENSE"],
    executables=executables,
    options={"build_exe": build_exe_options},
    install_requires=requirements
)

# NOTE:
# Rode o comando abaixo para gerar o executável:
# python .\setup.py build
