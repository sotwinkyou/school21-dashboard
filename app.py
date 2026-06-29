import streamlit as st
import pandas as pd
import requests
import urllib.parse
from docx import Document
from icalendar import Calendar
from datetime import datetime
import io
import time

# Настройки страницы Streamlit
st.set_page_config(
    page_title="Дашборд Школы 21", 
    layout="wide", 
    page_icon="💚"
)

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

# Кешируем запросы к Яндекс.Диску ровно на 2 минуты (120 секунд). 
# По истечении этого времени Streamlit автоматически сделает новый запрос к Яндексу, чтобы забрать свежие данные.
@st.cache_data(ttl=120)
def download_from_yandex(public_url):
    """
    Использует публичное API Яндекс.Диска для получения прямой ссылки на скачивание файла.
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
                return None, f"Не удалось скачать файл. Код ошибки: {file_response.status_code}"
        else:
            return None, "Не удалось получить доступ к файлу. Проверьте, что ссылка публичная."
    except Exception as e:
        return None, f"Ошибка подключения к Яндекс.Диску: {str(e)}"

def parse_excel(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        return df, None
    except Exception as e:
        return None, f"Ошибка Excel: {str(e)}"

def parse_docx(file_bytes):
    try:
        doc = Document(io.BytesIO(file_bytes))
        data = []
        for table in doc.tables:
            for row in table.rows[1:]: 
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

# Считываем ссылки из адресной строки (чтобы они не стирались при обновлении страницы)
saved_url_1 = st.query_params.get("table", "")
saved_url_2 = st.query_params.get("calendar", "")

# --- БОКОВАЯ ПАНЕЛЬ ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/e/e0/Sberbank_Logo_2020.svg", width=120)
st.sidebar.header("📁 Импорт данных")

# Выбор источника данных
source_type = st.sidebar.radio("Выберите источник данных:", ["Яндекс.Диск ссылки", "Локальные файлы"])

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
    st.sidebar.info("💡 Убедитесь, что ваши ссылки публичные. После ввода ссылок добавьте страницу в Закладки, чтобы не вставлять их повторно!")
    
    # Текстовые поля с авто-заполнением из адресной строки
    yandex_url_1 = st.sidebar.text_input(
        "Ссылка на Таблицу (Excel) или Документ (Word):", 
        value=saved_url_1,
        placeholder="https://disk.yandex.ru/i/..."
    )
    yandex_url_2 = st.sidebar.text_input(
        "Ссылка на Календарь (.ics) [опционально]:", 
        value=saved_url_2,
        placeholder="https://disk.yandex.ru/d/..."
    )
    
    # Сохраняем новые ссылки в адресную строку браузера при изменении
    if yandex_url_1 != saved_url_1:
        st.query_params["table"] = yandex_url_1
    if yandex_url_2 != saved_url_2:
        st.query_params["calendar"] = yandex_url_2

    # Кнопка для мгновенного принудительного обновления (сброс кеша)
    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Синхронизировать сейчас"):
        st.cache_data.clear()  # Полностью очищаем кеш
        st.toast("Кэш очищен! Загружаем самые свежие версии с Яндекс.Диска...", icon="📥")
        time.sleep(1)
        st.rerun()

    urls_to_process = [u for u in [yandex_url_1, yandex_url_2] if u]
    
    if urls_to_process:
        with st.sidebar.spinner("Подключение к Яндекс.Диску..."):
            for idx, url in enumerate(urls_to_process):
                file_bytes, err = download_from_yandex(url)
                if err:
                    errors.append(f"Ошибка загрузки (Ссылка #{idx+1}): {err}")
                elif file_bytes:
                    # Пытаемся автоматически определить тип по структуре / ссылке
                    if ".xlsx" in url or "excel" in url.lower() or idx == 0 and not url.endswith('.docx'):
                        df, err = parse_excel(file_bytes)
                    elif ".docx" in url or "word" in url.lower():
                        df, err = parse_docx(file_bytes)
                    elif ".ics" in url or "calendar" in url.lower():
                        df, err = parse_ical(file_bytes)
                    else:
                        df, err = parse_excel(file_bytes)
                        if err:
                            df, err = parse_docx(file_bytes)
                    
                    if err:
                        errors.append(f"Не удалось распознать формат по ссылке #{idx+1}: {err}")
                    elif df is not None:
                        all_events = pd.concat([all_events, df], ignore_index=True)

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
    st.metric(label="📅 Мероприятия", value=len(all_events) if not all_events.empty else "12")
    st.markdown('</div>', unsafe_allow_html=True)
with kpi_col4:
    st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
    st.metric(label="⏳ Ближайший дедлайн", value="2 дня", delta="Финал Бассейна", delta_color="inverse")
    st.markdown('</div>', unsafe_allow_html=True)

# Информационный блок о времени последнего обновления
current_time = datetime.now().strftime("%H:%M:%S")
st.caption(f"⏱️ Автоматическое обновление данных с Яндекс.Диска включено (каждые 2 минуты). Последняя синхронизация страницы: **{current_time}**.")

st.markdown("---")

# 2. РАСПИСАНИЕ И КАЛЕНДАРЬ
st.subheader("🗓️ Расписание и график событий")

if not all_events.empty:
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
        "Time": ["10:00", "14:00", "16:30", "12:00"],
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
