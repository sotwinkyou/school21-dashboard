import streamlit as st
import pandas as pd
import requests
import urllib.parse
from docx import Document
from icalendar import Calendar
from datetime import datetime
import io
import time
import zipfile

# --- ПРЕДОТВРАЩЕНИЕ ОШИБОК СТИЛЕЙ (MONKEYPATCH OPENPYXL) ---
try:
    import openpyxl.reader.excel
    orig_read_style = openpyxl.reader.excel.ExcelReader.read_style
    def safe_read_style(self):
        try:
            orig_read_style(self)
        except Exception:
            # Безопасные дефолты при критической ошибке XML-стилей Яндекса
            self.shared_styles = []
            self.stylesheet = None
    openpyxl.reader.excel.ExcelReader.read_style = safe_read_style
except Exception:
    pass

import openpyxl

# Настройки страницы Streamlit
st.set_page_config(
    page_title="Дашборд Школы 21", 
    layout="wide", 
    page_icon="💚"
)

# Очистка и запуск автоматического обновления (каждые 120 секунд)
if "last_update" not in st.session_state:
    st.session_state.last_update = time.time()

current_time = time.time()
if current_time - st.session_state.last_update > 120:
    st.session_state.last_update = current_time
    st.cache_data.clear()

# Фирменный стиль Сбера и Школы 21 (бело-зеленая гамма)
st.markdown("""
    <style>
    /* Основной фон приложения */
    .stApp {
        background-color: #f4f6f8;
    }
    
    /* Стилизация заголовков */
    h1 {
        color: #08a652 !important;
        font-family: 'SB Sans Text', 'Helvetica Neue', sans-serif;
        font-weight: 800;
        border-bottom: 3px solid #08a652;
        padding-bottom: 10px;
    }
    
    h2, h3 {
        color: #111827;
        font-family: 'SB Sans Text', sans-serif;
        font-weight: 600;
    }
    
    /* Стилизация карточек KPI */
    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: bold;
        color: #08a652 !important;
    }
    
    div[data-testid="stMetricLabel"] {
        font-size: 0.95rem !important;
        color: #4b5563 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Боковая панель */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e5e7eb;
    }
    
    /* Кнопки и интерактивные элементы */
    .stButton>button {
        background-color: #08a652 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton>button:hover {
        background-color: #06803f !important;
        box-shadow: 0 4px 12px rgba(8, 166, 82, 0.2);
    }
    
    /* Красивые контейнеры для KPI */
    .kpi-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid #08a652;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("💚 Автоматизированный дашборд | Школа 21")

# --- ФУНКЦИЯ ЗАГРУЗКИ С ЯНДЕКС.ДИСКА ---
@st.cache_data(ttl=120)
def download_from_yandex(public_url):
    """
    Скачивание файла по публичной ссылке Яндекс.Диска
    """
    base_url = 'https://cloud-api.yandex.net/v1/disk/public/resources/download'
    final_url = base_url + '?public_key=' + urllib.parse.quote(public_url)
    try:
        response = requests.get(final_url)
        if response.status_code == 200:
            download_url = response.json().get('href')
            file_response = requests.get(download_url)
            if file_response.status_code == 200:
                return file_response.content, None
            else:
                return None, f"Не удалось скачать файл. Код: {file_response.status_code}"
        else:
            return None, "Нет доступа. Проверьте, что ссылка публичная."
    except Exception as e:
        return None, f"Ошибка подключения к Яндексу: {str(e)}"

# --- ХИРУРГИЧЕСКОЕ УДАЛЕНИЕ СТИЛЕЙ ИЗ XLSX-АРХИВА ---
def strip_styles_from_xlsx(file_bytes):
    """
    Удаляет файл стилей xl/styles.xml из zip-архива xlsx.
    Это на 100% решает проблему с поломанными XML-стилями Яндекса.
    """
    try:
        in_mem_zip = io.BytesIO(file_bytes)
        out_mem_zip = io.BytesIO()
        with zipfile.ZipFile(in_mem_zip, 'r') as yin:
            with zipfile.ZipFile(out_mem_zip, 'w', zipfile.ZIP_DEFLATED) as yout:
                for item in yin.infolist():
                    # Полностью вырезаем стили, провоцирующие падение openpyxl
                    if item.filename == 'xl/styles.xml':
                        continue
                    yout.writestr(item, yin.read(item.filename))
        return out_mem_zip.getvalue()
    except Exception:
        return file_bytes

# --- ФУНКЦИИ УМНОГО ПАРСИНГА ФАЙЛОВ ---
def parse_excel(file_bytes):
    """
    Парсинг Excel с предварительной очисткой от проблемных стилей
    """
    try:
        # Очищаем байты файла от XML-стилей Яндекса перед загрузкой
        sanitized_bytes = strip_styles_from_xlsx(file_bytes)
        
        # Загружаем очищенную версию таблицы в безопасном read_only режиме
        wb = openpyxl.load_workbook(io.BytesIO(sanitized_bytes), read_only=True, data_only=True)
        sheet = wb.active
        
        data = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                data.append(list(row))
        
        if not data:
            return None, "Файл пуст или не содержит данных."
            
        # Формируем DataFrame
        headers = [str(h).strip() if h is not None else f"Column_{i}" for i, h in enumerate(data[0])]
        df = pd.DataFrame(data[1:], columns=headers)
        
        # Умный маппинг колонок (игнорирует регистр букв и пробелы)
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if "назван" in col_lower or "событи" in col_lower or "мероприяти" in col_lower:
                column_mapping[col] = "Название"
            elif "дат" in col_lower:
                column_mapping[col] = "Дата"
            elif "врем" in col_lower or "час" in col_lower:
                column_mapping[col] = "Время"
            elif "мест" in col_lower or "аудитор" in col_lower or "кластер" in col_lower:
                column_mapping[col] = "Место"
            elif "участн" in col_lower or "кол-во" in col_lower or "люд" in col_lower:
                column_mapping[col] = "Участники"
                
        df = df.rename(columns=column_mapping)
        
        # Убедимся, что все 5 колонок присутствуют в результирующей таблице
        required_cols = ["Название", "Дата", "Время", "Место", "Участники"]
        for rc in required_cols:
            if rc not in df.columns:
                df[rc] = "Не указано"
                
        return df[required_cols], None
    except Exception as e:
        return None, f"Ошибка Excel: {str(e)}"

def parse_docx(file_bytes):
    try:
        doc = Document(io.BytesIO(file_bytes))
        data = []
        for table in doc.tables:
            for row in table.rows[1:]: # пропускаем шапку таблицы
                text = [cell.text.strip() for cell in row.cells]
                if len(text) >= 5:
                    data.append(text[:5])
        df = pd.DataFrame(data, columns=["Название", "Дата", "Время", "Место", "Участники"])
        return df, None
    except Exception as e:
        return None, f"Ошибка Word: {str(e)}"

def parse_ical(file_bytes):
    try:
        gcal = Calendar.from_ical(file_bytes)
        data = []
        for component in gcal.walk():
            if component.name == "VEVENT":
                summary = str(component.get('summary'))
                start = component.get('dtstart').dt
                location = str(component.get('location', 'Кампус'))
                
                date_str = start.strftime("%Y-%m-%d") if isinstance(start, datetime) else str(start)
                time_str = start.strftime("%H:%M") if isinstance(start, datetime) else "00:00"
                
                data.append([summary, date_str, time_str, location, "Не указано"])
        df = pd.DataFrame(data, columns=["Название", "Дата", "Время", "Место", "Участники"])
        return df, None
    except Exception as e:
        return None, f"Ошибка Календаря: {str(e)}"

# --- БОКОВАЯ ПАНЕЛЬ И НАСТРОЙКИ СВЯЗИ ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/e/e0/Sberbank_Logo_2020.svg", width=120)
st.sidebar.header("📁 Параметры импорта")

# Чтение ссылок из URL-адреса для перманентного сохранения состояния
query_params = st.query_params
default_table_url = query_params.get("table", "")
default_cal_url = query_params.get("cal", "")

source_type = st.sidebar.radio("Источник данных:", ["Яндекс.Диск ссылки", "Локальные файлы"])

all_events = pd.DataFrame(columns=["Название", "Дата", "Время", "Место", "Участники"])
errors = []

if source_type == "Локальные файлы":
    uploaded_files = st.sidebar.file_uploader(
        "Загрузите файлы (Excel, Word, iCal)", 
        type=["xlsx", "docx", "ics"], 
        accept_multiple_files=True
    )
    if uploaded_files:
        for file in uploaded_files:
            file_bytes = file.read()
            if file.name.endswith('.xlsx'):
                df, err = parse_excel(file_bytes)
            elif file.name.endswith('.docx'):
                df, err = parse_docx(file_bytes)
            elif file.name.endswith('.ics'):
                df, err = parse_ical(file_bytes)
            
            if err:
                errors.append(f"Ошибка в файле {file.name}: {err}")
            elif df is not None:
                all_events = pd.concat([all_events, df], ignore_index=True)

else:
    st.sidebar.info("💡 Данные будут автоматически синхронизироваться раз в 2 минуты.")
    yandex_url_1 = st.sidebar.text_input("Ссылка на Таблицу (Excel) или Документ (Word):", value=default_table_url, placeholder="https://disk.yandex.ru/i/...")
    yandex_url_2 = st.sidebar.text_input("Ссылка на Календарь (.ics) [опционально]:", value=default_cal_url, placeholder="https://disk.yandex.ru/d/...")
    
    # Сохраняем ссылки в параметры запроса браузера
    if yandex_url_1 or yandex_url_2:
        st.query_params["table"] = yandex_url_1
        st.query_params["cal"] = yandex_url_2

    urls_to_process = []
    if yandex_url_1: urls_to_process.append((yandex_url_1, "table"))
    if yandex_url_2: urls_to_process.append((yandex_url_2, "cal"))
    
    if urls_to_process:
        for url, u_type in urls_to_process:
            file_bytes, err = download_from_yandex(url)
            if err:
                errors.append(f"Ошибка загрузки по ссылке ({u_type}): {err}")
            elif file_bytes:
                if ".xlsx" in url or "excel" in url.lower() or u_type == "table":
                    df, err = parse_excel(file_bytes)
                elif ".docx" in url or "word" in url.lower():
                    df, err = parse_docx(file_bytes)
                elif ".ics" in url or "calendar" in url.lower() or u_type == "cal":
                    df, err = parse_ical(file_bytes)
                else:
                    df, err = parse_excel(file_bytes)
                
                if err:
                    errors.append(f"Ошибка распознавания данных: {err}")
                elif df is not None:
                    all_events = pd.concat([all_events, df], ignore_index=True)

    if st.sidebar.button("🔄 Синхронизировать сейчас"):
        st.cache_data.clear()
        st.rerun()

# --- ГЛАВНЫЙ ЭКРАН ---

# 1. БЛОК KPI
st.subheader("🚀 Пульс Кампуса")
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

with kpi_col1:
    st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
    st.metric(label="👥 Студенты (Сheck-in)", value="432", delta="+12 сегодня")
    st.markdown('</div>', unsafe_allow_html=True)
with kpi_col2:
    st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
    st.metric(label="💻 Активные проекты", value="18", delta="Ветка С & DevOps")
    st.markdown('</div>', unsafe_allow_html=True)
with kpi_col3:
    st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
    st.metric(label="📅 Мероприятия в сети", value=len(all_events) if not all_events.empty else "0")
    st.markdown('</div>', unsafe_allow_html=True)
with kpi_col4:
    st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
    st.metric(label="⏳ Ближайший дедлайн", value="2 дня", delta="Финал Бассейна", delta_color="inverse")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# 2. РАСПИСАНИЕ И КАЛЕНДАРЬ
st.subheader("🗓️ Расписание и график событий")

if not all_events.empty:
    # Удаляем пустые значения в столбце 'Место' для фильтра
    all_events['Место'] = all_events['Место'].fillna('Не указано')
    places = all_events['Место'].unique()
    selected_place = st.multiselect("Фильтр по локациям (кластерам):", options=places, default=places)
    
    filtered_events = all_events[all_events['Место'].isin(selected_place)]
    st.dataframe(filtered_events, use_container_width=True)
else:
    st.info("Используйте боковую панель для загрузки локальных файлов или подключения ссылок Яндекс.Диска.")
    st.caption("Пример отображения расписания (демо-данные):")
    demo_data = pd.DataFrame({
        "Название": ["Хакатон Сбера", "Peer-to-Peer Защиты", "Встреча с HR", "Лекция по AI"],
        "Дата": ["2026-07-01", "2026-07-02", "2026-07-03", "2026-07-04"],
        "Время": ["10:00", "14:00", "16:30", "12:00"],
        "Место": ["Конференц-зал", "Кластер А", "Переговорка 2", "Кластер Б"],
        "Участники": ["120 человек", "Все пиры", "Команда Core", "Бассейн"]
    })
    st.dataframe(demo_data, use_container_width=True)

st.markdown("---")

# 3. ОТЧЕТЫ И БЛОКЕРЫ
st.subheader("📋 Оперативные отчеты")
col_left, col_right = st.columns(2)

with col_left:
    st.write("**📊 Прогресс подготовки к ивентам:**")
    st.progress(0.8, text="Организация Летнего Хакатона (80%)")
    st.progress(0.4, text="Закупка мерча для бассейна (40%)")

with col_right:
    st.write("**⚠️ Контроль блокеров и импорта:**")
    if errors:
        for err in errors:
            st.error(err)
    else:
        st.success("Все импортированные файлы синхронизированы без ошибок.")
