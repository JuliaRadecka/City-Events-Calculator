import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(layout="wide")

EXCEL_FILE = "Калькулятор_оценки_мероприятий_по_городам.xlsx"
TEMPLATE_SHEET = "TEMPLATE"
SHEET_CITIES = "ЦА по городам"

def to_json_safe_df(df: pd.DataFrame) -> pd.DataFrame:
    # Critical fix: column names from header=None are often numpy.int64 -> not JSON serializable for st-aggrid.
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    df = df.where(pd.notnull(df), "")

    def conv(x):
        if isinstance(x, np.generic):
            return x.item()
        return x

    return df.applymap(conv)

@st.cache_data
def load_template_df() -> pd.DataFrame:
    raw = pd.read_excel(EXCEL_FILE, sheet_name=TEMPLATE_SHEET, header=None)
    return to_json_safe_df(raw)

@st.cache_data
def load_geo_list() -> list:
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CITIES)
    vals = df.iloc[:, 0].dropna().astype(str).tolist()
    seen = set()
    out = []
    for v in vals:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out

st.title("Калькулятор оценки мероприятий")

geo_list = load_geo_list()
geo = st.selectbox("Гео (ввод)", geo_list, index=0 if geo_list else 0)

st.markdown("---")
st.subheader("Параметры")

c1, c2, c3, c4 = st.columns([2, 2, 1.5, 2])
with c1:
    ca = st.text_input("ЦА (авто, будет рассчитано)", value="", disabled=True)
with c2:
    venue_type = st.selectbox("Тип площадки (ввод)", ["Площадка", "Концерт", "Фестиваль"])
with c3:
    venue_days = st.number_input("Количество дней (ввод)", min_value=0, step=1, value=0)
with c4:
    period_days = st.number_input("Общий период размещения (дней) (ввод)", min_value=0, step=1, value=0)

visitors_plan = st.number_input("План посетителей (тыс. чел.) (ввод)", min_value=0.0, step=0.1, value=0.0)

st.info("Поля «(ввод)» заполняет пользователь. Поля «(авто)» заполняются после нажатия «Рассчитать».")

st.markdown("---")
st.subheader("Таблица (как в Excel)")

df = load_template_df()

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, filter=False, sortable=False, resizable=True)

# Temporary mapping: read-only (blue) columns by index from header=None
READ_ONLY_COLS = ["7", "8", "9", "10", "14", "15"]

for col_name in READ_ONLY_COLS:
    if col_name in df.columns:
        gb.configure_column(
            col_name,
            editable=False,
            cellStyle={"backgroundColor": "#1f4fd8", "color": "white"}
        )

# Editable columns grey (to show what user fills)
for col_name in df.columns:
    if col_name not in READ_ONLY_COLS:
        gb.configure_column(
            col_name,
            editable=True,
            cellStyle={"backgroundColor": "#d9d9d9", "color": "black"}
        )

grid_response = AgGrid(
    df,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.VALUE_CHANGED,
    height=650,
    theme="streamlit",
    allow_unsafe_jscode=True,
)

st.markdown("---")
if st.button("Рассчитать"):
    result_df = grid_response["data"].copy()

    # Placeholder calc for v2: fill auto columns so user sees the behavior
    for col_name in READ_ONLY_COLS:
        if col_name in result_df.columns:
            result_df[col_name] = "CALC"

    st.success("Расчёт выполнен (v2: Excel-grid). Следующий шаг — реальная математика и точная разметка серых/синих ячеек из TEMPLATE.")

    AgGrid(
        result_df,
        gridOptions=gb.build(),
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True,
    )
