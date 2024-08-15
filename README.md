# App de extração de cotações e geração de relatórios

Este repositório contém um script Python que automatiza a coleta de dados de 
cotação do dólar americano (USD) em relação ao real brasileiro (BRL) de um site 
específico, gera um relatório em PDF com a data, cotação, nome do site e 
captura de tela da página, e cria um instalador para o programa.

## Funcionalidades

- Possui instalador do programa
- Possui arquivo executável
- Extrai cotações de um site específico
- Gera um relatório em formato Word com:
    - Data e hora da extração
    - URL do site
    - Cotação extraída
    - Screenshot da página com a cotação
    - Autor do relatório
- Salva o relatório em uma pasta específica
- Converte o arquivo Word para PDF

### Relatório sendo gerado:

[![Assista ao vídeo no YouTube](https://img.youtube.com/vi/RgUZGl4Ke4w/0.jpg)](https://www.youtube.com/watch?v=RgUZGl4Ke4w)

## Requisitos

- Python 3.7 ou superior
- Ter instalado `MS Word` ou `LibreOffice` (Caso contrário, será instalado o 
  `Pandoc` no sitema automaticamente para executar a conversão do arquivo 
  `.docx` para PDF)
- Bibliotecas:

|   Acesso à Web   | Gerenciamento de Arquivos | Manipulação de Dados | Formatação e Conversão |
| :--------------: | :-----------------------: | :------------------: | :--------------------: |
|    `selenium`    |           `os`            |        `json`        |          `re`          |
| `urllib.request` |        `tempfile`         |      `datetime`      |         `docx`         |
|                  |         `shutil`          |        `time`        |      `subprocess`      |
|                  |         `pathlib`         |                      |                        |

## Como usar

### Download do Instalador:

O instalador esta disponível 
[aqui](https://github.com/eliasalbuquerque/python-scraper-and-report-generator/blob/main/download/mysetup.exe), 
faça o download e inicie a instalação.

### Modo desenvolvedor:

1. **Clonar o repositório:**

```bash
git clone https://github.com/eliasalbuquerque/python-scraper-and-report-generator
```

2. **Instalar as dependências:**

```bash
pip install -r requirements.txt
```

3. **Executar o script:**

```bash
python app.py
```

#### Gerar o executável da aplicação:


```bash
 python .\setup.py build
```

O executável `app.exe` estará na pasta `.\build\exe.win`.

Após executar o script, o app acessará o site definido no arquivo `config.json`, 
extrairá a cotação, gerará o relatório em MS Word e em PDF e o salvará na pasta 
`reports`.

## Notas

- **Edição do módulo `report.py`:** É necessário editar o módulo `report.py` 
  para personalizar o conteúdo do relatório.
- **Configuração do arquivo `config.json`:** É necessário configurar o arquivo 
  `config.json` para usar outros websites.


## Contribuições

O software está na primeira versão e precisa de melhorias para uma melhor 
experiência do usuário. Algumas sugestões:

- **Interface Gráfica:** Implementar uma interface gráfica para acompanhar o 
  progresso da geração do relatório.
- **Entrada de dados:** Adicionar a possibilidade de inserir dados como:
    - Nome do usuário
    - Nome do relatório
    - Pasta de destino do relatório

Agradeço a sua contribuição! 
