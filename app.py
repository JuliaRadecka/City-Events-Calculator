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
geo = st.selectbox("Гео", cities)

st.markdown("---")
st.subheader("Параметры")

c1, c2, c3, c4 = st.columns(4)
with c1:
    ca = st.text_input("ЦА", value="", disabled=True)
with c2:
    venue_type = st.selectbox("Тип площадки", ["Площадка", "Концерт", "Фестиваль"])
with c3:
    venue_days = st.number_input("Количество дней", min_value=0, step=1)
with c4:
    period_days = st.number_input("Общий период размещения (дней)", min_value=0, step=1)

visitors_plan = st.number_input("План посетителей (тыс. чел.)", min_value=0.0, step=0.1)

st.markdown("---")
st.subheader("Таблица (как в Excel)")

df = load_template().fillna("")

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, filter=False, sortable=False, resizable=True)

READ_ONLY_COLS = [7, 8, 9, 10, 14, 15]

for col_idx in READ_ONLY_COLS:
    if col_idx < len(df.columns):
        gb.configure_column(
            df.columns[col_idx],
            editable=False,
            cellStyle={
                "backgroundColor": "#1f4fd8",
                "color": "white"
            }
        )

grid_response = AgGrid(
    df,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.VALUE_CHANGED,
    height=650,
    theme="streamlit",
    allow_unsafe_jscode=True
)

st.markdown("---")
if st.button("Рассчитать"):
    result_df = grid_response["data"].copy()

    for col_idx in READ_ONLY_COLS:
        if col_idx < len(result_df.columns):
            result_df.iloc[:, col_idx] = "CALC"

    st.success("Расчёт выполнен")

    AgGrid(
        result_df,
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True
    )
