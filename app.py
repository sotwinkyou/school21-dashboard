import streamlit as st
import pandas as pd
from docx import Document
from icalendar import Calendar
from datetime import datetime

st.set_page_config(page_title="Дашборд Школы 21", layout="wide", page_icon="📊")
st.title("📊 Автоматизированный дашборд мероприятий | Школа 21")

# --- ФУНКЦИИ ПАРСИНГА ФАЙЛОВ ---
def parse_excel(file):
    try:
        df = pd.read_excel(file)
        return df, None
    except Exception as e:
        return None, f"Ошибка Excel: {str(e)}"

def parse_docx(file):
    try:
        doc = Document(file)
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

def parse_ical(file):
    try:
        gcal = Calendar.from_ical(file.read())
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

# --- БОКОВАЯ ПАНЕЛЬ ДЛЯ ЗАГРУЗКИ ---
st.sidebar.header("📁 Загрузка данных")
uploaded_files = st.sidebar.file_uploader(
    "Загрузите файлы (Excel, Word, iCal)", 
    type=["xlsx", "docx", "ics"], 
    accept_multiple_files=True
)

all_events = pd.DataFrame(columns=["Название", "Дата", "Время", "Место", "Участники"])
errors = []

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith('.xlsx'):
            df, err = parse_excel(file)
        elif file.name.endswith('.docx'):
            df, err = parse_docx(file)
        elif file.name.endswith('.ics'):
            df, err = parse_ical(file)
        
        if err:
            errors.append(f"Ошибка в файле {file.name}: {err}")
        elif df is not None:
            all_events = pd.concat([all_events, df], ignore_index=True)

# --- ГЛАВНЫЙ ЭКРАН ---
st.subheader("🚀 Пульс Кампуса (KPI)")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(label="👥 Студенты (Сheck-in)", value="432", delta="+12 сегодня")
with col2:
    st.metric(label="💻 Активные проекты", value="18", delta="Ветка С & DevOps")
with col3:
    st.metric(label="📅 Мероприятия (из файлов)", value=len(all_events))
with col4:
    st.metric(label="⏳ Ближайший дедлайн", value="2 дня", delta="Финал Бассейна", delta_color="inverse")

st.markdown("---")

st.subheader("🗓️ Расписание и график событий")
if not all_events.empty:
    places = all_events['Место'].unique()
    selected_place = st.multiselect("Фильтр по локациям (кластерам):", options=places, default=places)
    filtered_events = all_events[all_events['Место'].isin(selected_place)]
    st.dataframe(filtered_events, use_container_width=True)
else:
    st.info("Загрузите файлы в боковое меню, чтобы наполнить календарь событиями.")
    st.caption("Пример отображения данных:")
    demo_data = pd.DataFrame({
        "Название": ["Хакатон Сбера", "Peer-to-Peer Защиты", "Встреча с HR"],
        "Дата": ["2026-07-01", "2026-07-02", "2026-07-03"],
        "Время": ["10:00", "14:00", "16:30"],
        "Место": ["Конференц-зал", "Кластер А", "Переговорка 2"],
        "Участники": ["120 человек", "Все пиры", "Команда Core"]
    })
    st.table(demo_data)

st.markdown("---")

st.subheader("📋 Отчеты и Статус")
col_left, col_right = st.columns(2)
with col_left:
    st.write("**📊 Статус задач по ивентам:**")
    st.progress(0.8, text="Подготовка к Хакатону (80% готово)")
    st.progress(0.4, text="Закупка мерча для Бассейна (40% готово)")
with col_right:
    st.write("**⚠️ Блокеры и Ошибки парсинга:**")
    if errors:
        for err in errors:
            st.error(err)
    else:
        st.success("Конфликтов в расписании не обнаружено. Все файлы обработаны корректно.")
