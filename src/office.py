
# TRANSFORMAR EM CLASSE: classe Office():



import re
import docx


def create_word_document(doc, filename='New document', doc_path=r'.'):
    """Cria um documento do Word com o conteúdo especificado.
    Args:
        filename (str, optional): Nome do arquivo do documento. Padrão 'New document'.
        doc_path (str, optional): Caminho do diretório para salvar o documento. Padrão (diretório atual).
    Returns:
        docx.Document: O objeto Document criado.
    """

    if not re.match(r'^[a-zA-Z0-9_\.\-]+$', filename):
        raise ValueError(
            "Nome de arquivo inválido. Utilize apenas letras, números, '_' e '-'.")
    if not os.path.isdir(doc_path):
        raise ValueError("O caminho do documento é inválido.")

    try:
        add_content(doc)
        file_path = os.path.join(doc_path, filename + '.docx')
        doc.save(file_path)
    except Exception as e:
        logging.error(f'Erro ao criar documento: {e}')
        return None

    return doc


def add_content(doc):
    """Adiciona conteúdo no documento do Word.
    Args:
        doc (docx.Document): O objeto Document a ser modificado.
    Raises:
        Exception: Se ocorrer um erro ao adicionar o conteúdo ao documento.
    """

    doc.add_heading('Título Principal', level=1)
    doc.add_paragraph('Este é o primeiro parágrafo.')
    doc.add_heading('Subtítulo', level=2)
    doc.add_paragraph('Este é o segundo parágrafo.')
