import tracemalloc
import streamlit as st
import os
from translate_script import (
    extract_text, translate_text_google, translate_text_marian, translate_text_openai,
    create_translation_table_markdown, extract_text_from_url, 
    create_table_with_styles, extract_text_from_docx, extract_text_from_pdf, 
    generate_docx, apply_styles_to_docx, apply_styles_directly
)
from transformers import MarianMTModel, MarianTokenizer
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from dotenv import load_dotenv

# Налаштування тимчасової директорії
TEMP_DIR = "temp"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Завантаження змінних середовища
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")

# Перевірка наявності API-ключа
if not openai_api_key:
    st.error("Не знайдено OpenAI API ключ. Перевірте файл .env.")
    st.stop()

# Ініціалізація MarianMT
model_name = "Helsinki-NLP/opus-mt-en-uk"
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)

# Налаштування Streamlit
st.set_page_config(page_title="LegalTransUA", layout="wide")

# Підключення стилів
with open("static/css/style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Додавання зображення
st.image("static/img/LegalTransUA.jpg", use_container_width=True)

# Меню навігації
st.sidebar.title("Меню навігації")
section = st.sidebar.radio(
    "Перейдіть до розділу:",
    ["Головна сторінка", "Про додаток", "Контакти", "Допомога ЗСУ", "Корисні посилання"]
)

# Функція для збереження завантаженого файлу
def save_uploaded_file(uploaded_file):
    file_path = os.path.join(TEMP_DIR, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

# Функція обробки перекладу
def process_translation(paragraphs, base_name):
    google_translations = [""] * len(paragraphs)
    marian_translations = [""] * len(paragraphs)
    openai_translations = [""] * len(paragraphs)

    # Прогрес бари
    st.write("Прогрес перекладу Google Translate:")
    google_progress = st.progress(0)
    st.write("Прогрес перекладу MarianMT:")
    marian_progress = st.progress(0)
    st.write("Прогрес перекладу OpenAI GPT:")
    openai_progress = st.progress(0)

    # Виконання перекладів у потоках
    with ThreadPoolExecutor(max_workers=5) as executor:
        google_futures = {executor.submit(translate_text_google, para): idx for idx, para in enumerate(paragraphs)}
        for i, future in enumerate(as_completed(google_futures)):
            idx = google_futures[future]
            google_translations[idx] = future.result() or "Помилка перекладу"
            google_progress.progress((i + 1) / len(paragraphs))

        marian_futures = {executor.submit(translate_text_marian, para, tokenizer, model): idx for idx, para in enumerate(paragraphs)}
        for i, future in enumerate(as_completed(marian_futures)):
            idx = marian_futures[future]
            marian_translations[idx] = future.result() or "Помилка перекладу"
            marian_progress.progress((i + 1) / len(paragraphs))

        openai_futures = {executor.submit(translate_text_openai, para): idx for idx, para in enumerate(paragraphs)}
        for i, future in enumerate(as_completed(openai_futures)):
            idx = openai_futures[future]
            openai_translations[idx] = future.result() or "Помилка перекладу"
            openai_progress.progress((i + 1) / len(paragraphs))

    # Перевірка
    if all(not para for para in google_translations + marian_translations + openai_translations):
        logging.error("Усі переклади порожні. Документ не буде створено.")
        st.error("Переклад не виконався. Будь ласка, перевірте введений текст або джерело.")
        return False

    # Генерація Markdown-файлу
    markdown_content = create_translation_table_markdown(paragraphs, google_translations, marian_translations, openai_translations)
    markdown_file = os.path.join(TEMP_DIR, f"{base_name}_Translated.md")
    logging.info(f"Створення Markdown-файлу: {markdown_file}")

    with open(markdown_file, "w", encoding="utf-8") as f:
        f.write(markdown_content)

    # Конвертація в DOCX
    output_file = os.path.join(TEMP_DIR, f"{base_name}_Translated.docx")
    pandoc_command = f'pandoc -f markdown -t docx "{markdown_file}" -o "{output_file}"'
    conversion_result = os.system(pandoc_command)
    if conversion_result != 0:
        logging.error("Помилка при конвертації Markdown в DOCX.")
        st.error("Не вдалося створити DOCX-файл.")
        return False

    # Застосування стилів через apply_styles_directly
    styled_file = apply_styles_directly(output_file)
    if not styled_file:
        st.error("Не вдалося застосувати стилі до документа.")
        return False

    # Вивантаження файлу
    st.success("Переклад завершено!")
    st.download_button(
        label="Завантажити таблицю DOCX",
        data=open(styled_file, "rb").read(),
        file_name=os.path.basename(styled_file),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return True

# Головна логіка
if section == "Головна сторінка":
    st.title("LegalTransUA")
    st.header("Перекладач документів")
    st.write("Завантажте файл (DOCX або PDF) або введіть URL для перекладу.")

    type_of_source = st.radio("Оберіть тип джерела:", ["Файл", "URL"])

    if type_of_source == "Файл":
        uploaded_file = st.file_uploader("Завантажте файл (DOCX або PDF):", type=["docx", "pdf"])
        if uploaded_file:
            file_path = save_uploaded_file(uploaded_file)
            base_name, ext = os.path.splitext(uploaded_file.name)
            if ext.lower() not in [".docx", ".pdf"]:
                st.error("Підтримуються лише файли у форматах DOCX або PDF.")
            else:
                st.success(f"Файл '{uploaded_file.name}' успішно завантажено.")
                if st.button("Розпочати переклад"):
                    paragraphs = extract_text(file_path)
                    if paragraphs:
                        process_translation(paragraphs, base_name)
                    else:
                        st.error("Не вдалося отримати текст із документа.")

    elif type_of_source == "URL":
        url = st.text_input("Введіть URL:")
        if url and st.button("Розпочати переклад"):
            paragraphs = extract_text_from_url(url)
            if paragraphs:
                process_translation(paragraphs, "URL_translation")
            else:
                st.warning("Не вдалося знайти текст на сторінці.")

elif section == "Про додаток":
    st.title("Про LegalTransUA")
    st.write("**LegalTransUA** — це інноваційний додаток для автоматизації перекладу юридичних документів.")
    st.write("### Основні можливості:")
    st.write("- Переклад тексту з англійської на українську.")
    st.write("- Генерація таблиць із перекладом.")
    st.write("- Інтеграція із сучасними AI-інструментами.")

elif section == "Допомога ЗСУ":
    st.title("Підтримайте ЗСУ!")
    st.write("Ми вдячні нашим захисникам за можливість працювати у незалежній Україні.")
    st.write("Рекомендуємо підтримати фонд 'Повернись живим':")
    st.write("[Повернись живим](https://savelife.in.ua/)")

elif section == "Корисні посилання":
    st.title("Корисні посилання")
    st.write("- [Офіс ефективного регулювання (BRDO)](https://brdo.com.ua/)")
    st.write("- [Верховна Рада України](https://www.rada.gov.ua/)")
    st.write("- [Президент України](https://www.president.gov.ua/)")