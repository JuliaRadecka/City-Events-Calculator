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

# ---------------- Header ----------------
st.title("📊 Калькулятор оценки мероприятий")

# ---------------- Filters (compact, 1/4 width) ----------------
FILTER_W = 250  # px
geo_list = load_geo_list()

left, _right = st.columns([1, 3], gap="large")

with left:
    # 1) CA (auto, blue) – keep as input for familiarity, but readonly so user can't edit
    st.text_input(
        "ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)",
        value=st.session_state.get("ca_field", ""),
        key="ca_field",
    )

    # 2) Geo
    st.selectbox("Гео", geo_list, index=0 if geo_list else 0, key="geo_field")

    # 3) Venue type + days (same row, inside compact column)
    c1, c2 = st.columns([1, 1], gap="small")
    with c1:
        st.selectbox("Тип площадки", ["Площадка", "Концерт", "Фестиваль"], index=0, key="venue_type_field")
    with c2:
        st.number_input("Количество дней на фестивале/площадке", min_value=0, step=1, value=0, key="venue_days_field")

    # 4) Visitors plan
    st.number_input("План посетителей (в тысячах человек)", min_value=0.0, step=0.1, value=0.0, key="visitors_plan_field")

# ---------------- Force Excel-like colors + fixed widths (robust) ----------------
# Reason: Streamlit's DOM doesn't reliably keep widgets inside our wrapper div,
# so CSS classes may not apply. We target widgets by aria-label (their label) via JS.
components.html(
    f"""
    <script>
      const root = window.parent.document;

      function styleWidget(label, bg, fg, bold=false, readonly=false) {{
        const el = root.querySelector(`[aria-label="${{label}}"]`);
        if (!el) return;

        // Width (apply to input/select and also to visible control wrappers)
        el.style.width = "{FILTER_W}px";
        el.style.maxWidth = "{FILTER_W}px";

        // Base styling
        el.style.background = bg;
        el.style.color = fg;
        if (bold) el.style.fontWeight = "600";

        // Readonly for CA
        if (readonly) {{
          el.setAttribute("readonly", "true");
          el.style.cursor = "not-allowed";
        }}

        // Try to also style the nearest visible container (selectbox uses baseweb wrappers)
        let p = el.closest('div[data-baseweb="select"]');
        if (p) {{
          p.style.width = "{FILTER_W}px";
          p.style.maxWidth = "{FILTER_W}px";
          const inner = p.querySelector('div');
          if (inner) {{
            inner.style.background = bg;
            inner.style.color = fg;
          }}
        }}

        // Streamlit widget containers
        let w = el.closest('.stTextInput, .stNumberInput, .stSelectbox');
        if (w) {{
          w.style.width = "{FILTER_W}px";
          w.style.maxWidth = "{FILTER_W}px";
        }}
      }}

      // Apply styles (Excel semantics)
      styleWidget("ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)", "#1f4fd8", "#ffffff", true, true);
      styleWidget("Гео", "#d9d9d9", "#111111", false, false);
      styleWidget("Тип площадки", "#d9d9d9", "#111111", false, false);
      styleWidget("Количество дней на фестивале/площадке", "#d9d9d9", "#111111", false, false);
      styleWidget("План посетителей (в тысячах человек)", "#d9d9d9", "#111111", false, false);

      // Also shrink the whole left column feel by reducing padding around widgets
      const widgets = root.querySelectorAll(".stTextInput, .stNumberInput, .stSelectbox");
      widgets.forEach(w => {{ w.style.marginBottom = "10px"; }});
    </script>
    """,
    height=0,
)

st.markdown("---")

# ---------------- Table (unchanged) ----------------
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
            cellStyle={"backgroundColor": "#1f4fd8", "color": "white"},
        )

for col_name in df.columns:
    if col_name not in READ_ONLY_COLS:
        gb.configure_column(
            col_name,
            editable=True,
            cellStyle={"backgroundColor": "#d9d9d9", "color": "black"},
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
    edited = pd.DataFrame(grid_response["data"]).copy()

    for col_name in READ_ONLY_COLS:
        if col_name in edited.columns:
            edited[col_name] = "CALC"

    # Placeholder: show that CA gets populated after calculation
    st.session_state["ca_field"] = "CALC"

    st.success("Расчёт выполнен (на этом шаге правим только фильтры: порядок/цвета/компактность).")

    AgGrid(
        edited,
        gridOptions=gb.build(),
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True,
    )
