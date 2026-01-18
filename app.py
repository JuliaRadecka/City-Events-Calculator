import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(layout="wide")

EXCEL_FILE = "Калькулятор_оценки_мероприятий_по_городам.xlsx"
TEMPLATE_SHEET = "TEMPLATE"
SHEET_CITIES = "ЦА по городам"

def to_json_safe_df(df: pd.DataFrame) -> pd.DataFrame:
    # Ensure JSON-serializable grid: column names as strings and numpy scalars -> python scalars
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

# ---------- Styling helpers (compact + color-coding like Excel) ----------
st.markdown(
    """<style>
    .filters-wrap {max-width: 820px;}
    .auto-blue input, .auto-blue textarea {
        background-color: #1f4fd8 !important;
        color: white !important;
    }
    .user-grey input, .user-grey textarea {
        background-color: #d9d9d9 !important;
        color: black !important;
    }
    .user-grey div[data-baseweb="select"] > div {
        background-color: #d9d9d9 !important;
        color: black !important;
    }
    </style>""",
    unsafe_allow_html=True
)

# ---------- UI ----------
st.title("📊 Калькулятор оценки мероприятий")

geo_list = load_geo_list()

st.markdown('<div class="filters-wrap">', unsafe_allow_html=True)

# 1) CA (auto, blue, after calculate)
st.markdown('<div class="auto-blue">', unsafe_allow_html=True)
ca_value = st.text_input(
    "ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)",
    value="",
    disabled=True
)
st.markdown('</div>', unsafe_allow_html=True)

# 2) GEO (user, grey)
st.markdown('<div class="user-grey">', unsafe_allow_html=True)
geo = st.selectbox("Гео", geo_list, index=0 if geo_list else 0)
st.markdown('</div>', unsafe_allow_html=True)

# 3) Venue type + days (same row)
c1, c2 = st.columns([2, 1])
with c1:
    st.markdown('<div class="user-grey">', unsafe_allow_html=True)
    venue_type = st.selectbox("Тип площадки", ["Площадка", "Концерт", "Фестиваль"], index=0)
    st.markdown('</div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="user-grey">', unsafe_allow_html=True)
    venue_days = st.number_input("Количество дней на фестивале/площадке", min_value=0, step=1, value=0)
    st.markdown('</div>', unsafe_allow_html=True)

# 4) Visitors plan (next row)
st.markdown('<div class="user-grey">', unsafe_allow_html=True)
visitors_plan = st.number_input("План посетителей (в тысячах человек)", min_value=0.0, step=0.1, value=0.0)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# Table renders below filters (as requested)
st.subheader("Таблица (как в Excel)")
df = load_template_df()

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, filter=False, sortable=False, resizable=True)

# Temporary mapping for blue calculated columns (will later be derived from TEMPLATE fills)
READ_ONLY_COLS = ["7", "8", "9", "10", "14", "15"]
for col_name in READ_ONLY_COLS:
    if col_name in df.columns:
        gb.configure_column(
            col_name,
            editable=False,
            cellStyle={"backgroundColor": "#1f4fd8", "color": "white"}
        )

# Grey for all other columns (input)
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
    edited = pd.DataFrame(grid_response["data"])

    # Placeholder: fill blue columns to show the behavior
    for col_name in READ_ONLY_COLS:
        if col_name in edited.columns:
            edited[col_name] = "CALC"

    st.success("Расчёт выполнен (на этом шаге обновляем блок фильтров и демонстрируем заполнение авто-полей).")

    AgGrid(
        edited,
        gridOptions=gb.build(),
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True,
    )
