import os
import logging
import zipfile
import xml.etree.ElementTree as ET
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
from transformers import MarianMTModel, MarianTokenizer
from deep_translator import GoogleTranslator
import openai
import requests
from bs4 import BeautifulSoup
import fitz  # PyMuPDF
import re
import time
from datetime import datetime
import shutil
from lxml import etree
import subprocess

# Ваші інші імпорти і змінні тут

__all__ = [
    "extract_text_from_docx",
    "extract_text_from_pdf",
    "extract_text_from_url",
    "translate_text_google",
    "translate_text_marian",
    "translate_text_openai",
    "create_translation_table_markdown",
    "generate_docx",
    "apply_styles_to_docx",  # Додайте сюди цю функцію
]

# Налаштування логування
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Завантаження змінних середовища
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    logging.error("API-ключ OpenAI не знайдено. Перевірте файл .env.")
    exit(1)
openai.api_key = openai_api_key

# Ініціалізація MarianMT
model_name = "Helsinki-NLP/opus-mt-en-uk"
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)

ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

# -------------------- Допоміжні функції --------------------

def sanitize_text(text):
    return re.sub(r'[^ -~]', '', text)

def sanitize_text_for_xml(text):
    """Очищає текст для використання в XML."""
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def sanitize_filename(filename):
    name, ext = os.path.splitext(filename)
    return re.sub(r'[<>:"/\\|?*]', '_', name) + ext

# -------------------- Екстракція тексту --------------------

def extract_text_from_docx(file_path):
    """Витягує текст із DOCX-файлу."""
    with zipfile.ZipFile(file_path, 'r') as docx:
        xml_content = docx.read('word/document.xml')
        tree = ET.ElementTree(ET.fromstring(xml_content))
        root = tree.getroot()
        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        paragraphs = [node.text for node in root.findall('.//w:t', namespaces)]
    return paragraphs

def extract_text_from_pdf(file_path):
    """Витягує текст із PDF-файлу."""
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"
    return [line.strip() for line in text.splitlines() if line.strip()]

def extract_text_from_url(url):
    """Витягує текст із веб-сторінки."""
    response = requests.get(url)
    if response.status_code != 200:
        raise Exception(f"Не вдалося завантажити сторінку: {url}")
    soup = BeautifulSoup(response.content, "html.parser")
    paragraphs = soup.find_all("p")
    return [para.get_text().strip() for para in paragraphs if para.get_text().strip()]

def extract_text(source):
    try:
        if source.startswith("http"):
            logging.info("Source identified as URL.")
            return extract_text_from_url(source)
        elif source.lower().endswith(".pdf"):
            logging.info("Source identified as PDF.")
            return extract_text_from_pdf(source)
        elif source.lower().endswith(".docx"):
            logging.info("Source identified as DOCX.")
            return extract_text_from_docx(source)
        else:
            raise ValueError("Unsupported file format. Supported formats: DOCX, PDF, or URL.")
    except Exception as e:
        logging.error(f"Error extracting text: {e}")
        return []

# -------------------- Переклад тексту --------------------

def translate_text_google(text):
    try:
        return GoogleTranslator(source='en', target='uk').translate(text)
    except Exception as e:
        logging.error(f"Google Translate Error: {e}")
        return "Translation error"

def translate_text_marian(text, tokenizer, model):
    try:
        inputs = tokenizer([text], return_tensors="pt", padding=True, truncation=True)
        translated = model.generate(**inputs)
        return tokenizer.batch_decode(translated, skip_special_tokens=True)[0]
    except Exception as e:
        logging.warning(f"MarianMT Error: {e}")
        return "Translation error"

def translate_text_openai(text, max_retries=3):
    for attempt in range(max_retries):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Translate the following text to Ukrainian."},
                    {"role": "user", "content": text},
                ],
            )
            return response.choices[0].message["content"].strip()
        except Exception as e:
            logging.warning(f"OpenAI Error (attempt {attempt + 1}/{max_retries}): {e}")
            time.sleep(2 ** attempt + 1)
    return "Translation error"

def get_default_content_types():
    """Повертає стандартний XML для [Content_Types].xml."""
    return """<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
        <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
    </Types>
    """

def get_relationships():
    """Повертає XML для _rels/.rels."""
    return """<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>
    """

def get_default_styles():
    """Повертає стандартний XML для стилів."""
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Normal">
            <w:name w:val="Normal"/>
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>
                <w:sz w:val="20"/> <!-- Розмір 10pt -->
            </w:rPr>
        </w:style>
        <w:style w:type="table" w:styleId="TableGrid">
            <w:name w:val="Table Grid"/>
            <w:tblPr>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
        </w:style>
    </w:styles>
    """

def get_document_rels():
    """Повертає XML для word/_rels/document.xml.rels."""
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
    </Relationships>
    """

def get_numbering_xml():
    """Повертає XML для word/numbering.xml."""
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
            <w:lvl w:ilvl="0">
                <w:start w:val="1"/>
                <w:numFmt w:val="decimal"/>
                <w:lvlText w:val="%1."/>
                <w:lvlJc w:val="left"/>
                <w:pPr>
                    <w:ind w:left="720" w:hanging="360"/>
                </w:pPr>
            </w:lvl>
        </w:abstractNum>
        <w:num w:numId="1">
            <w:abstractNumId w:val="0"/>
        </w:num>
    </w:numbering>
    """

def verify_zip_content(file_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            invalid_files = z.testzip()
            if invalid_files is None:
                logging.info(f"Файл {file_path} успішно перевірено.")
            else:
                logging.warning(f"Некоректні файли: {invalid_files}")
    except Exception as e:
        logging.error(f"Помилка перевірки ZIP: {e}")

# -------------------- Формування документа --------------------

def validate_and_write_xml(zipfile_obj, file_name, xml_content):
    try:
        if isinstance(xml_content, str):
            xml_content = xml_content.encode('utf-8')

        etree.fromstring(xml_content)  # Перевірка валідності XML
        zipfile_obj.writestr(file_name, xml_content)
        logging.info(f"Файл {file_name} успішно додано до архіву.")
    except etree.XMLSyntaxError as e:
        logging.error(f"Помилка у {file_name}: {e}")
        raise

def create_translation_table(root, paragraphs, google_translations, marian_translations, openai_translations):
    """Додає таблицю перекладів до документа."""
    body = root.find(".//w:body", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    if body is None:
        body = ET.SubElement(root, "w:body")

    table = ET.SubElement(body, "w:tbl")

    headers = ["Оригінальний текст", "Google Translate", "MarianMT", "OpenAI GPT"]
    header_row = ET.SubElement(table, "w:tr")
    for header in headers:
        header_cell = ET.SubElement(header_row, "w:tc")
        header_text = ET.SubElement(header_cell, "w:t")
        header_text.text = sanitize_text_for_xml(header)

    for para, g_trans, m_trans, o_trans in zip(paragraphs, google_translations, marian_translations, openai_translations):
        row = ET.SubElement(table, "w:tr")
        for text in [para, g_trans, m_trans, o_trans]:
            cell = ET.SubElement(row, "w:tc")
            text_element = ET.SubElement(cell, "w:t")
            text_element.text = sanitize_text_for_xml(text)

def create_translation_table_markdown(paragraphs, google_translations, marian_translations, openai_translations):
    """
    Створює таблицю у форматі Markdown з оригінальним текстом та перекладами.
    """
    header = (
        "# Automated Document Translation\n\n"
        "Generated using the **LegalTransUA** script.\n"
        f"Date and time of translation: **{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}**.\n\n"
    )
    table_header = "| No | Original Text | Google Translate | MarianMT | OpenAI GPT |\n"
    table_divider = "|:---|:------------------|:----------------|:---------|:----------|\n"
    rows = [
        f"| {i+1} | {para} | {google or '-'} | {marian or '-'} | {openai or '-'} |"
        for i, (para, google, marian, openai) in enumerate(zip(paragraphs, google_translations, marian_translations, openai_translations))
    ]
    return header + table_header + table_divider + "\n".join(rows)

def create_table_with_styles(data):
    """Генерує XML таблиці зі стилями."""
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    etree.register_namespace('w', namespaces['w'])

    table = etree.Element(f"{{{namespaces['w']}}}tbl")
    tbl_pr = etree.SubElement(table, f"{{{namespaces['w']}}}tblPr")
    etree.SubElement(tbl_pr, f"{{{namespaces['w']}}}tblStyle", {f"{{{namespaces['w']}}}val": "TableGrid"})

    # Заголовок таблиці
    header_row = etree.SubElement(table, f"{{{namespaces['w']}}}tr")
    for header_text in ["#", "Оригінал", "Google Translate", "MarianMT", "OpenAI GPT"]:
        cell = etree.SubElement(header_row, f"{{{namespaces['w']}}}tc")
        text = etree.SubElement(cell, f"{{{namespaces['w']}}}p")
        run = etree.SubElement(text, f"{{{namespaces['w']}}}r")
        run_pr = etree.SubElement(run, f"{{{namespaces['w']}}}rPr")
        etree.SubElement(run_pr, f"{{{namespaces['w']}}}rFonts", {
            f"{{{namespaces['w']}}}ascii": "Arial",
            f"{{{namespaces['w']}}}hAnsi": "Arial",
            f"{{{namespaces['w']}}}eastAsia": "Arial",
            f"{{{namespaces['w']}}}cs": "Arial"
        })
        etree.SubElement(run_pr, f"{{{namespaces['w']}}}sz", {f"{{{namespaces['w']}}}val": "20"})
        etree.SubElement(run, f"{{{namespaces['w']}}}t").text = header_text

    # Дані таблиці
    for i, row_data in enumerate(data, start=1):
        row = etree.SubElement(table, f"{{{namespaces['w']}}}tr")
        for cell_data in [str(i)] + list(row_data):  # Перетворення row_data на список
            cell = etree.SubElement(row, f"{{{namespaces['w']}}}tc")
            text = etree.SubElement(cell, f"{{{namespaces['w']}}}p")
            run = etree.SubElement(text, f"{{{namespaces['w']}}}r")
            run_pr = etree.SubElement(run, f"{{{namespaces['w']}}}rPr")
            etree.SubElement(run_pr, f"{{{namespaces['w']}}}rFonts", {
                f"{{{namespaces['w']}}}ascii": "Arial",
                f"{{{namespaces['w']}}}hAnsi": "Arial",
                f"{{{namespaces['w']}}}eastAsia": "Arial",
                f"{{{namespaces['w']}}}cs": "Arial"
            })
            etree.SubElement(run_pr, f"{{{namespaces['w']}}}sz", {f"{{{namespaces['w']}}}val": "20"})
            etree.SubElement(run, f"{{{namespaces['w']}}}t").text = cell_data

    return table

def validate_xml_structure(xml_content):
    try:
        etree.fromstring(xml_content.encode('utf-8'))  # Перетворення на байти
        logging.info("XML структура валідна.")
    except etree.XMLSyntaxError as e:
        logging.error(f"Помилка у структурі XML: {e}")
        raise

def verify_docx_integrity(docx_file):
    """
    Перевіряє цілісність DOCX-файлу.
    """
    try:
        with zipfile.ZipFile(docx_file, 'a') as docx:
            docx.writestr('word/footer1.xml', generate_footer_with_page_numbers())
            docx.writestr('word/_rels/document.xml.rels', get_document_rels_with_footer())
            invalid_files = [name for name in docx.namelist() if not docx.testzip()]
            if invalid_files:
                logging.warning(f"Некоректні файли: {invalid_files}")
            else:
                logging.info(f"Файл {docx_file} валідний.")
    except Exception as e:
        logging.error(f"Помилка перевірки DOCX: {e}")
        raise

def apply_table_styles(docx_file):
    try:
        with zipfile.ZipFile(docx_file, 'r') as docx:
            document_xml = docx.read('word/document.xml')
            styles_xml = docx.read('word/styles.xml')

        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        doc_tree = etree.fromstring(document_xml)
        styles_tree = etree.fromstring(styles_xml)

        # Перевірка та оновлення стилю TableGrid
        table_grid_style = styles_tree.find(".//w:style[@w:styleId='TableGrid']", namespaces)
        if table_grid_style is None:
            table_grid_style = etree.SubElement(styles_tree, f"{{{namespaces['w']}}}style", {
                f"{{{namespaces['w']}}}type": "table",
                f"{{{namespaces['w']}}}styleId": "TableGrid"
            })
            etree.SubElement(table_grid_style, f"{{{namespaces['w']}}}name", {
                f"{{{namespaces['w']}}}val": "Table Grid"
            })
            tbl_pr = etree.SubElement(table_grid_style, f"{{{namespaces['w']}}}tblPr")
            borders = etree.SubElement(tbl_pr, f"{{{namespaces['w']}}}tblBorders")
            for border in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                etree.SubElement(borders, f"{{{namespaces['w']}}}{border}", {
                    f"{{{namespaces['w']}}}val": "single",
                    f"{{{namespaces['w']}}}sz": "4",
                    f"{{{namespaces['w']}}}color": "000000"
                })

        # Застосування стилю до таблиць
        for table in doc_tree.findall('.//w:tbl', namespaces):
            tbl_pr = table.find('w:tblPr', namespaces)
            if tbl_pr is None:
                tbl_pr = etree.SubElement(table, f"{{{namespaces['w']}}}tblPr")
            tbl_style = tbl_pr.find('w:tblStyle', namespaces)
            if tbl_style is None:
                tbl_style = etree.SubElement(tbl_pr, f"{{{namespaces['w']}}}tblStyle")
            tbl_style.set(f"{{{namespaces['w']}}}val", "TableGrid")

        # Запис оновлених XML
        with zipfile.ZipFile(docx_file, 'w') as docx:
            docx.writestr('word/document.xml', etree.tostring(doc_tree, pretty_print=True, encoding='utf-8'))
            docx.writestr('word/styles.xml', etree.tostring(styles_tree, pretty_print=True, encoding='utf-8'))
        logging.info("Стилі таблиць успішно застосовані.")
    except Exception as e:
        logging.error(f"Помилка при застосуванні стилів таблиць: {e}")

apply_table_styles("path_to_your_file.docx")

def generate_footer_with_page_numbers():
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:ftr xmlns:w="{namespaces['w']}">
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
                <w:instrText xml:space="preserve"> PAGE </w:instrText>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
            <w:r>
                <w:t> з </w:t>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
                <w:instrText xml:space="preserve"> NUMPAGES </w:instrText>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
        </w:p>
    </w:ftr>
    """

def add_footer_to_docx(docx_file):
    try:
        footer_xml = generate_footer_with_page_numbers()
        with zipfile.ZipFile(docx_file, 'a') as docx:
            docx.writestr('word/footer1.xml', footer_xml)
            docx.writestr('word/_rels/document.xml.rels', get_document_rels_with_footer())
        logging.info("Нижній колонтитул з нумерацією сторінок додано.")
    except Exception as e:
        logging.error(f"Помилка при додаванні нижнього колонтитула: {e}")

def get_document_rels_with_footer():
    """Генерує XML для document.xml.rels з нижнім колонтитулом."""
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
    </Relationships>
    """

def generate_basic_docx(output_file):
    """Генерує базовий DOCX-файл."""
    document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            <w:p>
                <w:r>
                    <w:t>Hello, World!</w:t>
                </w:r>
            </w:p>
        </w:body>
    </w:document>
    """

    with zipfile.ZipFile(output_file, 'w') as docx:
        validate_and_write_xml(docx, '[Content_Types].xml', get_default_content_types())
        validate_and_write_xml(docx, '_rels/.rels', get_relationships())
        validate_and_write_xml(docx, 'word/document.xml', document_xml)
        validate_and_write_xml(docx, 'word/styles.xml', get_default_styles())
        validate_and_write_xml(docx, 'word/footer1.xml', generate_footer_with_page_numbers(namespaces))

def generate_docx_with_table(output_file):
    """Генерує DOCX-файл з простою таблицею."""
    table_xml = """<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:tr>
            <w:tc><w:t>Header 1</w:t></w:tc>
            <w:tc><w:t>Header 2</w:t></w:tc>
        </w:tr>
        <w:tr>
            <w:tc><w:t>Row 1, Cell 1</w:t></w:tc>
            <w:tc><w:t>Row 1, Cell 2</w:t></w:tc>
        </w:tr>
    </w:tbl>"""

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            {table_xml}
        </w:body>
    </w:document>
    """

    with zipfile.ZipFile(output_file, 'w') as docx:
        docx.writestr('[Content_Types].xml', get_default_content_types())
        docx.writestr('_rels/.rels', get_relationships())
        docx.writestr('word/document.xml', document_xml)
        docx.writestr('word/styles.xml', get_default_styles())

   

def generate_docx(output_file, table_xml):
    """Генерує DOCX-файл із таблицею."""
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="{namespaces['w']}">
        <w:body>
            {etree.tostring(table_xml, encoding="unicode")}
            <w:sectPr>
                <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
                <w:pgMar w:top="360" w:bottom="360" w:left="360" w:right="360"/>
            </w:sectPr>
        </w:body>
    </w:document>
    """
    with zipfile.ZipFile(output_file, 'w') as docx:
        docx.writestr('[Content_Types].xml', """<?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
        </Types>
        """)
        docx.writestr('_rels/.rels', """<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """)
        docx.writestr('word/document.xml', document_xml)
        logging.info(f"Документ збережено у файл: {output_file}")


def validate_docx_integrity(docx_file):
    try:
        with zipfile.ZipFile(docx_file, 'r') as docx:
            invalid_file = docx.testzip()
            if invalid_file is None:
                logging.info(f"Файл {docx_file} валідний.")
            else:
                logging.warning(f"Некоректний файл: {invalid_file}")
    except Exception as e:
        logging.error(f"Помилка перевірки DOCX: {e}")
        raise

def validate_xml(xml_content):
    """Перевіряє XML на валідність."""
    try:
        etree.fromstring(xml_content.encode('utf-8'))
        logging.info("XML валідний.")
    except etree.XMLSyntaxError as e:
        logging.error(f"Некоректний XML: {e}")
        raise

def apply_styles_to_docx(docx_file):
    try:
        with zipfile.ZipFile(docx_file, 'r') as docx:
            xml_content = docx.read('word/document.xml')
            root = etree.fromstring(xml_content)

        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # Застосування стилів до таблиць
        for table in root.findall('.//w:tbl', namespaces):
            tbl_pr = table.find('w:tblPr', namespaces)
            if tbl_pr is None:
                tbl_pr = etree.SubElement(table, f"{{{namespaces['w']}}}tblPr")
            tbl_style = tbl_pr.find('w:tblStyle', namespaces)
            if tbl_style is None:
                tbl_style = etree.SubElement(tbl_pr, f"{{{namespaces['w']}}}tblStyle")
                tbl_style.set(f"{{{namespaces['w']}}}val", "TableGrid")

        # Налаштування орієнтації сторінки та полів
        sect_pr = root.find('.//w:sectPr', namespaces)
        if sect_pr is None:
            body = root.find('.//w:body', namespaces)
            sect_pr = etree.SubElement(body, f"{{{namespaces['w']}}}sectPr")

        pg_sz = sect_pr.find('w:pgSz', namespaces)
        if pg_sz is None:
            pg_sz = etree.SubElement(sect_pr, f"{{{namespaces['w']}}}pgSz")
        pg_sz.set(f"{{{namespaces['w']}}}orient", "landscape")
        pg_sz.set(f"{{{namespaces['w']}}}w", "16838")  # Ширина
        pg_sz.set(f"{{{namespaces['w']}}}h", "11906")  # Висота

        pg_mar = sect_pr.find('w:pgMar', namespaces)
        if pg_mar is None:
            pg_mar = etree.SubElement(sect_pr, f"{{{namespaces['w']}}}pgMar")
        pg_mar.set(f"{{{namespaces['w']}}}top", "360")
        pg_mar.set(f"{{{namespaces['w']}}}bottom", "360")
        pg_mar.set(f"{{{namespaces['w']}}}left", "360")
        pg_mar.set(f"{{{namespaces['w']}}}right", "360")

        # Запис оновленого документа
        updated_content = etree.tostring(root, pretty_print=True, encoding='utf-8', xml_declaration=True)
        with zipfile.ZipFile(docx_file, 'w') as docx:
            docx.writestr('word/document.xml', updated_content)
            docx.writestr('[Content_Types].xml', get_default_content_types())
            docx.writestr('_rels/.rels', get_relationships())
            docx.writestr('word/styles.xml', get_default_styles())
        logging.info("Стилі успішно застосовані.")
    except Exception as e:
        logging.error(f"Помилка при застосуванні стилів: {e}")

def apply_styles_directly(docx_path):
    try:
        # Створення тимчасового каталогу для роботи з файлами
        temp_dir = "temp_docx"
        shutil.rmtree(temp_dir, ignore_errors=True)
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Шлях до основного XML документа
        document_xml_path = os.path.join(temp_dir, "word/document.xml")

        # Редагування стилів у document.xml
        with open(document_xml_path, 'r', encoding='utf-8') as file:
            document_xml = file.read()
        root = etree.fromstring(document_xml.encode('utf-8'))

        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # Додавання тексту перед таблицею
        body = root.find(".//w:body", namespaces)
        if body is None:
            body = etree.SubElement(root, f"{{{namespaces['w']}}}body")

        # Напис "Automated Document Translation"
        title_paragraph = etree.SubElement(body, f"{{{namespaces['w']}}}p")
        title_pPr = etree.SubElement(title_paragraph, f"{{{namespaces['w']}}}pPr")
        title_jc = etree.SubElement(title_pPr, f"{{{namespaces['w']}}}jc")
        title_jc.set(f"{{{namespaces['w']}}}val", "center")  # Вирівнювання по центру
        title_run = etree.SubElement(title_paragraph, f"{{{namespaces['w']}}}r")
        title_text = etree.SubElement(title_run, f"{{{namespaces['w']}}}t")
        title_text.text = "Automated Document Translation"

        # Напис з інформацією про генерацію
        info_paragraph = etree.SubElement(body, f"{{{namespaces['w']}}}p")
        info_run = etree.SubElement(info_paragraph, f"{{{namespaces['w']}}}r")
        info_text = etree.SubElement(info_run, f"{{{namespaces['w']}}}t")
        info_text.text = f"Файл згенеровано з використанням скрипту LegalTransUA\nДата і час генерації: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        info_text.set(f"{{{namespaces['w']}}}space", "preserve")  # Збереження перенесень рядків

        # Налаштування стилів таблиці
        for table in root.findall('.//w:tbl', namespaces):
            tbl_pr = table.find('w:tblPr', namespaces)
            if tbl_pr is None:
                tbl_pr = etree.SubElement(table, f"{{{namespaces['w']}}}tblPr")

            # Додавання чорних меж таблиці
            tbl_borders = tbl_pr.find('w:tblBorders', namespaces)
            if tbl_borders is None:
                tbl_borders = etree.SubElement(tbl_pr, f"{{{namespaces['w']}}}tblBorders")
            for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = tbl_borders.find(f'w:{side}', namespaces)
                if border is None:
                    border = etree.SubElement(tbl_borders, f"{{{namespaces['w']}}}{side}")
                border.set(f"{{{namespaces['w']}}}val", "single")
                border.set(f"{{{namespaces['w']}}}sz", "4")
                border.set(f"{{{namespaces['w']}}}color", "000000")

            # Задаємо ширину стовпців
            for row_idx, row in enumerate(table.findall('.//w:tr', namespaces)):
                for cell_idx, cell in enumerate(row.findall('.//w:tc', namespaces)):
                    # Встановлення ширини стовпців
                    tc_pr = cell.find('w:tcPr', namespaces)
                    if tc_pr is None:
                        tc_pr = etree.SubElement(cell, f"{{{namespaces['w']}}}tcPr")
                    tc_w = etree.SubElement(tc_pr, f"{{{namespaces['w']}}}tcW")
                    if cell_idx == 0:  # Перший стовпець
                        tc_w.set(f"{{{namespaces['w']}}}w", "400")  # Вузький стовпець
                    else:
                        tc_w.set(f"{{{namespaces['w']}}}w", "2500")  # Рівна ширина для інших
                    tc_w.set(f"{{{namespaces['w']}}}type", "dxa")

                    # Налаштування заливки та стилів
                    shd = tc_pr.find('w:shd', namespaces)
                    if shd is None:
                        shd = etree.SubElement(tc_pr, f"{{{namespaces['w']}}}shd")
                    if row_idx == 0:  # Заголовок таблиці
                        shd.set(f"{{{namespaces['w']}}}fill", "ADD8E6")  # Блакитний
                        # Напівжирний шрифт, вирівнювання по центру
                        p_pr = cell.find('w:p/w:pPr', namespaces)
                        if p_pr is None:
                            p_pr = etree.SubElement(cell.find('w:p', namespaces), f"{{{namespaces['w']}}}pPr")
                        jc = etree.SubElement(p_pr, f"{{{namespaces['w']}}}jc")
                        jc.set(f"{{{namespaces['w']}}}val", "center")
                        for r in cell.findall('.//w:r', namespaces):
                            r_pr = r.find('w:rPr', namespaces)
                            if r_pr is None:
                                r_pr = etree.SubElement(r, f"{{{namespaces['w']}}}rPr")
                            b = etree.SubElement(r_pr, f"{{{namespaces['w']}}}b")
                    elif cell_idx == 0:  # Стовпець порядкових номерів
                        shd.set(f"{{{namespaces['w']}}}fill", "D3D3D3")  # Сірий

                    # Вирівнювання тексту по ширині
                    p_pr = cell.find('w:p/w:pPr', namespaces)
                    if p_pr is None:
                        p_pr = etree.SubElement(cell.find('w:p', namespaces), f"{{{namespaces['w']}}}pPr")
                    jc = p_pr.find('w:jc', namespaces)
                    if jc is None:
                        jc = etree.SubElement(p_pr, f"{{{namespaces['w']}}}jc")
                    jc.set(f"{{{namespaces['w']}}}val", "both")  # Вирівнювання по ширині

                    # Arial шрифт, розмір 20% менший
                    for r in cell.findall('.//w:r', namespaces):
                        r_pr = r.find('w:rPr', namespaces)
                        if r_pr is None:
                            r_pr = etree.SubElement(r, f"{{{namespaces['w']}}}rPr")
                        r_fonts = etree.SubElement(r_pr, f"{{{namespaces['w']}}}rFonts")
                        r_fonts.set(f"{{{namespaces['w']}}}ascii", "Arial")
                        r_fonts.set(f"{{{namespaces['w']}}}hAnsi", "Arial")
                        sz = etree.SubElement(r_pr, f"{{{namespaces['w']}}}sz")
                        sz.set(f"{{{namespaces['w']}}}val", "16")  # Зменшення розміру шрифту на 20%

        # Налаштування горизонтальної орієнтації та вузьких полів
        sect_pr = root.find(".//w:sectPr", namespaces)
        if sect_pr is None:
            sect_pr = etree.SubElement(body, f"{{{namespaces['w']}}}sectPr")
        pg_sz = sect_pr.find("w:pgSz", namespaces)
        if pg_sz is None:
            pg_sz = etree.SubElement(sect_pr, f"{{{namespaces['w']}}}pgSz")
        pg_sz.set(f"{{{namespaces['w']}}}orient", "landscape")  # Горизонтальна орієнтація
        pg_sz.set(f"{{{namespaces['w']}}}w", "16838")  # Ширина
        pg_sz.set(f"{{{namespaces['w']}}}h", "11906")  # Висота
        pg_mar = sect_pr.find("w:pgMar", namespaces)
        if pg_mar is None:
            pg_mar = etree.SubElement(sect_pr, f"{{{namespaces['w']}}}pgMar")
        pg_mar.set(f"{{{namespaces['w']}}}top", "360")  # Верхнє поле
        pg_mar.set(f"{{{namespaces['w']}}}bottom", "360")  # Нижнє поле
        pg_mar.set(f"{{{namespaces['w']}}}left", "360")  # Ліве поле
        pg_mar.set(f"{{{namespaces['w']}}}right", "360")  # Праве поле

        # Збереження змін у document.xml
        with open(document_xml_path, 'wb') as file:
            file.write(etree.tostring(root, pretty_print=True, encoding='utf-8', xml_declaration=True))

        # Створення нового DOCX файлу
        updated_docx_path = docx_path.replace(".docx", "_Styled.docx")
        with zipfile.ZipFile(updated_docx_path, 'w') as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    arcname = os.path.relpath(filepath, temp_dir)
                    zip_ref.write(filepath, arcname)

        # Очищення тимчасових файлів
        shutil.rmtree(temp_dir)

        logging.info(f"Файл успішно збережено: {updated_docx_path}")
        return updated_docx_path
    except Exception as e:
        logging.error(f"Помилка при зміні стилів: {e}")
        return None


data = [
    ["Hello", "Привіт", "Привіт", "Привіт"],
    ["World", "Світ", "Світ", "Світ"],
    ["How are you?", "Як справи?", "Як справи?", "Як справи?"]
]

table_xml = create_table_with_styles(data)
output_file = "translation_table.docx"
generate_docx(output_file, table_xml)
apply_styles_directly(output_file)
