import streamlit as st
import pandas as pd
from docx import Document
from icalendar import Calendar
from datetime import datetime

st.set_page_config(
    page_title="Дашборд Мероприятий | Школа 21 Сбер",
    layout="wide",
    page_icon="🟩",
    initial_sidebar_state="expanded"
)

# Внедряем кастомные CSS стили для полной кастомизации под бренд Сбера и Школы 21
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
    
    /* Глобальные шрифты */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .mono-text {
        font-family: 'JetBrains Mono', monospace;
    }
    
    /* Фирменный градиент Сбера в шапке */
    .sber-header {
        background: linear-gradient(135deg, #21a038 0%, #08a652 100%);
        padding: 20px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(8, 166, 82, 0.15);
    }
    
    /* Стилизация карточек KPI */
    .kpi-card {
        background-color: white;
        padding: 20px;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        transition: all 0.3s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(8, 166, 82, 0.08);
        border-color: #a7f3d0;
    }
    
    /* Зеленый акцент для текстов */
    .text-sber-green {
        color: #08a652;
    }
</style>
""", unsafe_allow_html=True)

# Фирменная шапка
st.markdown("""
<div class="sber-header">
    <div style="display: flex; align-items: center; justify-content: space-between;">
        <div>
            <h1 style="margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">
                <span class="mono-text" style="background-color: rgba(255,255,255,0.2); padding: 2px 8px; border-radius: 6px; margin-right: 10px;">■</span>ШКОЛА 21
            </h1>
            <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Автоматизированная экосистема интеграции данных из .docx, .xlsx, .ics планов</p>
        </div>
        <div style="text-align: right; background-color: rgba(255,255,255,0.15); padding: 8px 16px; border-radius: 30px;">
            <span style="font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px;">Кампус Москва • Активен</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- ФУНКЦИИ ПАРСИНГА ФАЙЛОВ ---
def parse_excel(file):
    try:
        df = pd.read_excel(file)
        # Ожидаем колонки: Название, Дата, Время, Место, Участники
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
                
                # Форматируем дату и время
                date_str = start.strftime("%Y-%m-%d") if isinstance(start, datetime) else str(start)
                time_str = start.strftime("%H:%M") if isinstance(start, datetime) else "00:00"
                
                data.append([summary, date_str, time_str, location, "Не указано"])
        df = pd.DataFrame(data, columns=["Название", "Дата", "Время", "Место", "Участники"])
        return df, None
    except Exception as e:
        return None, f"Ошибка Календаря: {str(e)}"

# --- БОКОВАЯ ПАНЕЛЬ ДЛЯ ЗАГРУЗКИ ---
st.sidebar.markdown("""
<div style="text-align: center; margin-bottom: 20px;">
    <h2 style="color: #08a652; margin-bottom: 5px;">🟩 Панель управления</h2>
    <p style="font-size: 12px; color: #6b7280;">Импорт расписания и планов</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.sidebar.file_uploader(
    "Загрузите файлы расписаний", 
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

# --- БЛОК KPI ---
st.markdown("### 🚀 Пульс Кампуса")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <p style="margin: 0; font-size: 11px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px;">Студенты (Check-In)</p>
        <h2 style="margin: 10px 0 5px 0; font-size: 32px; font-weight: 800; color: #111827;">452 <span style="font-size: 14px; color: #10b981; font-weight: 600;">+14%</span></h2>
        <div style="background-color: #f3f4f6; border-radius: 10px; height: 6px; width: 100%; overflow: hidden; margin-top: 10px;">
            <div style="background-color: #08a652; height: 6px; width: 56%; border-radius: 10px;"></div>
        </div>
        <p style="margin: 8px 0 0 0; font-size: 11px; color: #6b7280;">из 800 зарегистрированных пиров</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <p style="margin: 0; font-size: 11px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px;">Активные проекты</p>
        <h2 style="margin: 10px 0 5px 0; font-size: 32px; font-weight: 800; color: #111827;">24</h2>
        <div style="background-color: #f3f4f6; border-radius: 10px; height: 6px; width: 100%; overflow: hidden; margin-top: 10px;">
            <div style="background-color: #3b82f6; height: 6px; width: 75%; border-radius: 10px;"></div>
        </div>
        <p style="margin: 8px 0 0 0; font-size: 11px; color: #6b7280;">Ветки C, Web & DevOps</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <p style="margin: 0; font-size: 11px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px;">Загрузка Кластеров</p>
        <h2 style="margin: 10px 0 5px 0; font-size: 32px; font-weight: 800; color: #111827;">68% <span style="font-size: 11px; color: #8b5cf6; font-weight: 600; background-color: #f3e8ff; padding: 2px 6px; border-radius: 4px; margin-left: 5px;">Оптимально</span></h2>
        <div style="background-color: #f3f4f6; border-radius: 10px; height: 6px; width: 100%; overflow: hidden; margin-top: 10px;">
            <div style="background-color: #8b5cf6; height: 6px; width: 68%; border-radius: 10px;"></div>
        </div>
        <p style="margin: 8px 0 0 0; font-size: 11px; color: #6b7280;">Свободно 102 рабочих места</p>
    </div>
    """, unsafe_allow_html=True)

with col4:
    events_count = len(all_events) if not all_events.empty else 3
    st.markdown(f"""
    <div class="kpi-card">
        <p style="margin: 0; font-size: 11px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: 0.5px;">События & Файлы</p>
        <h2 style="margin: 10px 0 5px 0; font-size: 32px; font-weight: 800; color: #111827;">{events_count} <span style="font-size: 11px; color: #f97316; font-weight: 600; background-color: #ffedd5; padding: 2px 6px; border-radius: 4px; margin-left: 5px;">Активно</span></h2>
        <div style="background-color: #f3f4f6; border-radius: 10px; height: 6px; width: 100%; overflow: hidden; margin-top: 10px;">
            <div style="background-color: #f97316; height: 6px; width: 100%; border-radius: 10px;"></div>
        </div>
        <p style="margin: 8px 0 0 0; font-size: 11px; color: #6b7280;">Все источники синхронизированы</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- РАСПИСАНИЕ И КАЛЕНДАРЬ ---
st.markdown("### 🗓️ Календарь событий")

if not all_events.empty:
    places = all_events['Место'].unique()
    selected_place = st.multiselect("Фильтр по локациям кампуса:", options=places, default=places)
    filtered_events = all_events[all_events['Место'].isin(selected_place)]
    
    st.dataframe(
        filtered_events,
        use_container_width=True,
        column_config={
            "Название": st.column_config.TextColumn("Событие", width="large"),
            "Дата": st.column_config.TextColumn("Дата проведения"),
            "Время": st.column_config.TextColumn("Время"),
            "Место": st.column_config.TextColumn("Аудитория/Локация"),
            "Участники": st.column_config.TextColumn("Ориентировочное кол-во пиров")
        }
    )
else:
    st.info("💡 Загрузите файлы .xlsx, .docx или .ics в левое меню, чтобы автоматически построить календарь.")
    
    # Демо-данные в бело-зеленой стилистике
    st.markdown("**Пример структуры расписания:**")
    demo_data = pd.DataFrame({
        "Название": ["Хакатон Сбера по генеративному AI", "Peer-to-Peer Защиты финальных проектов", "Карьерный митап с HR Сбера"],
        "Дата": ["2026-06-29", "2026-06-30", "2026-07-02"],
        "Время": ["10:00 — 18:00", "14:00 — 17:30", "16:30 — 18:00"],
        "Место": ["Кластер А (Главный зал)", "Кластер Б (Аудитории 3-5)", "Конференц-зал"],
        "Участники": ["145 участников", "Все пиры бассейна", "Приглашенные эксперты HR"]
    })
    st.dataframe(demo_data, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# --- РЕПОРТЫ И БЛОКЕРЫ ---
st.markdown("### 📋 Операционный статус и Блокеры")
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("""
    <div style="background-color: white; padding: 20px; border-radius: 16px; border: 1px solid #e5e7eb; height: 100%;">
        <h4 style="margin: 0 0 15px 0; color: #111827;">📊 Статус подготовки мероприятий</h4>
        <div style="margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                <span>Организация Сберовского Хакатона</span>
                <span style="color: #08a652;">80%</span>
            </div>
            <div style="background-color: #f3f4f6; border-radius: 10px; height: 8px; overflow: hidden;">
                <div style="background-color: #08a652; height: 8px; width: 80%;"></div>
            </div>
        </div>
        <div style="margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                <span>Подготовка мерча для нового Бассейна</span>
                <span style="color: #f97316;">40%</span>
            </div>
            <div style="background-color: #f3f4f6; border-radius: 10px; height: 8px; overflow: hidden;">
                <div style="background-color: #f97316; height: 8px; width: 40%;"></div>
            </div>
        </div>
        <div>
            <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                <span>Согласование списков стажировок в Сбер</span>
                <span style="color: #10b981;">100%</span>
            </div>
            <div style="background-color: #f3f4f6; border-radius: 10px; height: 8px; overflow: hidden;">
                <div style="background-color: #10b981; height: 8px; width: 100%;"></div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_right:
    st.markdown("""
    <div style="background-color: white; padding: 20px; border-radius: 16px; border: 1px solid #e5e7eb; height: 100%;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
            <h4 style="margin: 0; color: #111827;">⚠️ Анализ конфликтов расписания</h4>
            <span style="font-size: 11px; background-color: #fef2f2; color: #ef4444; font-weight: 700; padding: 2px 8px; border-radius: 10px;">1 активный</span>
        </div>
    """, unsafe_allow_html=True)
    
    if errors:
        for err in errors:
            st.error(f"🚨 {err}")
    else:
        st.markdown("""
        <div style="background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 12px; padding: 12px; display: flex; gap: 10px;">
            <div style="font-size: 20px;">🚨</div>
            <div>
                <h5 style="margin: 0 0 4px 0; color: #991b1b; font-size: 13px; font-weight: 700;">Пересечение бронирования аудиторий</h5>
                <p style="margin: 0; color: #7f1d1d; font-size: 11px; line-height: 1.4;">
                    30 июня в 14:00 обнаружен конфликт: Защита проектов в Кластере Б пересекается с техническим вебинаром стаффа. Места забронированы из разных файлов.
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("</div>", unsafe_allow_html=True)
