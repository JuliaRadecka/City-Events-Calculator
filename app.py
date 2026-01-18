import streamlit as st
import pandas as pd
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
import streamlit.components.v1 as components

st.set_page_config(layout="wide")

EXCEL_FILE = "Калькулятор_оценки_мероприятий_по_городам.xlsx"
TEMPLATE_SHEET = "TEMPLATE"
SHEET_CITIES = "ЦА по городам"

# ---------------- Data helpers ----------------
def to_json_safe_df(df: pd.DataFrame) -> pd.DataFrame:
    # st-aggrid: ensure JSON-serializable (column names as strings + numpy scalars -> python scalars)
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

# ---------------- UI styling (Excel-like filters) ----------------
FILTER_W = 250  # px per spec

st.markdown(
    f"""
    <style>
      /* Compact block (do not stretch full width) */
      .filters-wrap {{
        max-width: {FILTER_W*3 + 120}px;
        padding: 6px 0 2px 0;
      }}

      /* Force fixed widths for widgets */
      .filters-wrap .stTextInput,
      .filters-wrap .stNumberInput,
      .filters-wrap .stSelectbox {{
        width: {FILTER_W}px !important;
        max-width: {FILTER_W}px !important;
      }}

      .filters-wrap .stTextInput input,
      .filters-wrap .stNumberInput input {{
        width: {FILTER_W}px !important;
        max-width: {FILTER_W}px !important;
      }}

      .filters-wrap div[data-baseweb="select"] {{
        width: {FILTER_W}px !important;
        max-width: {FILTER_W}px !important;
      }}

      /* Grey (user input) */
      .user-grey input {{
        background: #d9d9d9 !important;
        color: #111 !important;
      }}
      .user-grey div[data-baseweb="select"] > div {{
        background: #d9d9d9 !important;
        color: #111 !important;
      }}

      /* Blue (auto) */
      .auto-blue input {{
        background: #1f4fd8 !important;
        color: #fff !important;
        font-weight: 600 !important;
      }}

      /* Make labels slightly tighter */
      .filters-wrap label {{
        margin-bottom: 2px !important;
      }}
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------- Header ----------------
st.title("📊 Калькулятор оценки мероприятий")

# ---------------- Filters block (order exactly as requested) ----------------
geo_list = load_geo_list()

st.markdown('<div class="filters-wrap">', unsafe_allow_html=True)

# 1) CA (auto, blue) – render as input but enforce readonly via JS (disabled breaks colors)
st.markdown('<div class="auto-blue">', unsafe_allow_html=True)
st.text_input(
    "ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)",
    value=st.session_state.get("ca_field", ""),
    key="ca_field"
)
st.markdown('</div>', unsafe_allow_html=True)

components.html(
    """
    <script>
      const root = window.parent.document;
      const inputs = root.querySelectorAll('input');
      for (const el of inputs) {
        if (el.getAttribute('aria-label') === 'ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)') {
          el.setAttribute('readonly', 'true');
          el.style.cursor = 'not-allowed';
        }
      }
    </script>
    """,
    height=0,
)

# 2) GEO (user, grey) – next line
st.markdown('<div class="user-grey">', unsafe_allow_html=True)
geo = st.selectbox("Гео", geo_list, index=0 if geo_list else 0, key="geo_field")
st.markdown('</div>', unsafe_allow_html=True)

# 3) Venue type + days on same line (both grey)
row3a, row3b = st.columns([1, 1], gap="large")
with row3a:
    st.markdown('<div class="user-grey">', unsafe_allow_html=True)
    venue_type = st.selectbox("Тип площадки", ["Площадка", "Концерт", "Фестиваль"], index=0, key="venue_type_field")
    st.markdown('</div>', unsafe_allow_html=True)
with row3b:
    st.markdown('<div class="user-grey">', unsafe_allow_html=True)
    venue_days = st.number_input("Количество дней на фестивале/площадке", min_value=0, step=1, value=0, key="venue_days_field")
    st.markdown('</div>', unsafe_allow_html=True)

# 4) Visitors plan – next line (grey)
st.markdown('<div class="user-grey">', unsafe_allow_html=True)
visitors_plan = st.number_input("План посетителей (в тысячах человек)", min_value=0.0, step=0.1, value=0.0, key="visitors_plan_field")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# ---------------- Table (kept as-is for this iteration) ----------------
st.subheader("Таблица (как в Excel)")

df = load_template_df()

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True, filter=False, sortable=False, resizable=True)

READ_ONLY_COLS = ["7", "8", "9", "10", "14", "15"]
for col_name in READ_ONLY_COLS:
    if col_name in df.columns:
        gb.configure_column(
            col_name,
            editable=False,
            cellStyle={"backgroundColor": "#1f4fd8", "color": "white"}
        )

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

    for col_name in READ_ONLY_COLS:
        if col_name in edited.columns:
            edited[col_name] = "CALC"

    # Placeholder: show that CA gets populated after calculation
    st.session_state["ca_field"] = "CALC"

    st.success("Расчёт выполнен (на этом шаге правим только фильтры: порядок/цвета/размеры).")

    AgGrid(
        edited,
        gridOptions=gb.build(),
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True,
    )
