import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta

# === НАСТРОЙКА СТРАНИЦЫ ===
st.set_page_config(
    page_title="Мой дашборд",
    page_icon="📊",
    layout="wide"
)

# === ЗАГОЛОВОК ===
st.title("📊 Интерактивный дашборд на Python")
st.markdown("---")

# === БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА ДАННЫХ ===
with st.sidebar:
    st.header("⚙️ Настройки")
    
    # Выбор источника данных
    source = st.radio(
        "Источник данных:",
        ["Сгенерировать пример", "Загрузить CSV-файл"]
    )
    
    if source == "Загрузить CSV-файл":
        uploaded_file = st.file_uploader(
            "Загрузите CSV-файл",
            type=["csv"],
            help="Файл должен содержать числовые колонки для графиков"
        )
        if uploaded_file is not None:
            df = pd.read_csv(uploaded_file)
            st.success(f"✅ Загружено {len(df)} строк")
        else:
            st.info("Загрузите CSV или используйте пример")
            df = None
    else:
        # Генерация примерных данных
        st.info("Генерируем демо-данные...")
        np.random.seed(42)
        dates = pd.date_range(start="2024-01-01", periods=365, freq="D")
        df = pd.DataFrame({
            "Дата": dates,
            "Продажи": np.random.randint(100, 500, 365) + np.sin(np.linspace(0, 20, 365)) * 50,
            "Посетители": np.random.randint(500, 2000, 365),
            "Категория": np.random.choice(["A", "B", "C", "D"], 365),
            "Регион": np.random.choice(["Север", "Юг", "Восток", "Запад"], 365)
        })
        df["Конверсия"] = (df["Продажи"] / df["Посетители"] * 100).round(2)
    
    st.markdown("---")
    st.caption(f"Данных: {len(df) if df is not None else 0} строк")

# === ПРОВЕРКА ДАННЫХ ===
if df is None or len(df) == 0:
    st.warning("⚠️ Нет данных для отображения. Загрузите файл или выберите генерацию примера.")
    st.stop()

# === ОСНОВНЫЕ ФИЛЬТРЫ ===
st.sidebar.subheader("🔍 Фильтры")

# Автоопределение колонок
numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

# Если есть дата-колонка, добавляем фильтр по дате
if date_cols:
    date_col = date_cols[0]
    min_date = df[date_col].min()
    max_date = df[date_col].max()
    date_range = st.sidebar.date_input(
        "Период",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    if len(date_range) == 2:
        start_date, end_date = date_range
        df = df[(df[date_col] >= pd.to_datetime(start_date)) & (df[date_col] <= pd.to_datetime(end_date))]

# Фильтры по категориальным колонкам
for col in categorical_cols[:3]:  # Ограничим 3 колонками
    unique_values = df[col].unique()
    selected = st.sidebar.multiselect(
        f"Фильтр по {col}",
        options=unique_values,
        default=unique_values
    )
    if selected:
        df = df[df[col].isin(selected)]

# === ОТОБРАЖЕНИЕ ДАННЫХ ===
if len(df) == 0:
    st.warning("⚠️ Нет данных после применения фильтров.")
    st.stop()

# === МЕТРИКИ (KPI) ===
st.subheader("📈 Ключевые показатели")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("📦 Всего записей", f"{len(df):,}")
with col2:
    if numeric_cols:
        total = df[numeric_cols[0]].sum()
        st.metric(f"💰 Сумма {numeric_cols[0]}", f"{total:,.0f}")
with col3:
    if numeric_cols:
        avg = df[numeric_cols[0]].mean()
        st.metric(f"📊 Среднее {numeric_cols[0]}", f"{avg:,.1f}")
with col4:
    if len(numeric_cols) > 1:
        corr = df[numeric_cols[:2]].corr().iloc[0, 1]
        st.metric("🔗 Корреляция", f"{corr:.2f}")

st.markdown("---")

# === ГРАФИКИ ===
col1, col2 = st.columns(2)

with col1:
    st.subheader("📉 Динамика")
    if date_cols:
        x_axis = date_cols[0]
    else:
        x_axis = df.index if "index" in df.columns else df.columns[0]
    
    if numeric_cols:
        fig_line = px.line(
            df,
            x=x_axis,
            y=numeric_cols[0],
            title=f"Тренд {numeric_cols[0]}",
            template="plotly_white"
        )
        st.plotly_chart(fig_line, use_container_width=True)

with col2:
    st.subheader("📊 Распределение")
    if categorical_cols:
        fig_bar = px.bar(
            df[categorical_cols[0]].value_counts().reset_index(),
            x="index",
            y=categorical_cols[0],
            title=f"Распределение по {categorical_cols[0]}",
            template="plotly_white"
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    elif numeric_cols:
        fig_hist = px.histogram(
            df,
            x=numeric_cols[0],
            title=f"Гистограмма {numeric_cols[0]}",
            template="plotly_white"
        )
        st.plotly_chart(fig_hist, use_container_width=True)

# === ДОПОЛНИТЕЛЬНЫЙ РЯД ГРАФИКОВ ===
st.subheader("📊 Дополнительный анализ")
col1, col2 = st.columns(2)

with col1:
    if categorical_cols and numeric_cols:
        fig_box = px.box(
            df,
            x=categorical_cols[0],
            y=numeric_cols[0],
            title=f"Разброс {numeric_cols[0]} по {categorical_cols[0]}",
            template="plotly_white"
        )
        st.plotly_chart(fig_box, use_container_width=True)

with col2:
    if len(numeric_cols) >= 2:
        fig_scatter = px.scatter(
            df,
            x=numeric_cols[0],
            y=numeric_cols[1],
            color=categorical_cols[0] if categorical_cols else None,
            title=f"Зависимость {numeric_cols[0]} от {numeric_cols[1]}",
            template="plotly_white"
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

# === ТАБЛИЦА ===
with st.expander("📋 Просмотр данных"):
    st.dataframe(df, use_container_width=True, height=400)
    
    # Кнопка скачивания
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Скачать CSV",
        data=csv,
        file_name=f"dashboard_data_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

st.caption("🚀 Дашборд создан на Python + Streamlit")
