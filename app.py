import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(layout="wide")

EXCEL_FILE = "Калькулятор_оценки_мероприятий_по_городам.xlsx"
TEMPLATE_SHEET = "TEMPLATE"

@st.cache_data
def load_template():
    return pd.read_excel(EXCEL_FILE, sheet_name=TEMPLATE_SHEET, header=None)

@st.cache_data
def load_cities():
    df = pd.read_excel(EXCEL_FILE, sheet_name="ЦА по городам")
    return df.iloc[:, 0].dropna().unique().tolist()

st.title("Калькулятор оценки мероприятий")

cities = load_cities()
city = st.selectbox("Город", cities)

st.markdown("---")
st.subheader("Параметры")

c1, c2, c3 = st.columns(3)
with c1:
    days_event = st.number_input("Дней мероприятия", min_value=0, step=1)
with c2:
    period_days = st.number_input("Период размещения (дней)", min_value=0, step=1)
with c3:
    visitors_plan = st.number_input("План посетителей", min_value=0.0, step=0.1)

st.markdown("---")
st.subheader("Таблица (как в Excel)")

df = load_template().fillna("")

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, filter=True, sortable=True, resizable=True)

READ_ONLY_COLS = [7, 8, 9, 10, 14, 15]

for col_idx in READ_ONLY_COLS:
    if col_idx < len(df.columns):
        gb.configure_column(
            df.columns[col_idx],
            editable=False,
            cellStyle={"backgroundColor": "#1f4fd8", "color": "white"}
        )

grid = AgGrid(
    df,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.VALUE_CHANGED,
    height=600,
    theme="streamlit"
)

st.markdown("---")
if st.button("Рассчитать"):
    result = grid["data"].copy()
    for col_idx in READ_ONLY_COLS:
        if col_idx < len(result.columns):
            result.iloc[:, col_idx] = "CALC"

    st.success("Расчёт выполнен (v2: Excel-grid)")

    AgGrid(result, height=600, theme="streamlit")
