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

# ---------------- Filters (compact, fixed widths) ----------------
FILTER_W = 250  # px (per spec)
USER_GREY = "#c7c7c7"  # darker in light theme
AUTO_BLUE = "#1f4fd8"

geo_list = load_geo_list()

# A slightly wider "left" column to keep long captions on one line,
# while input controls themselves remain 250px.
left, _right = st.columns([1.3, 3.7], gap="large")

# CSS: keep our custom captions on one line
st.markdown(
    """
    <style>
      .one-line-caption {
        white-space: nowrap;
        overflow: visible;
        font-size: 0.78rem;
        margin-bottom: 2px;
      }
      .one-line-caption strong { font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True
)

# Labels (full, one line) per your request
CA_LABEL = "ЦА (Унифицированная аудитория для всех медиа, тыс. 16+)"
DAYS_LABEL = "Количество дней на фестивале/площадке"

with left:
    # 1) CA: show full caption (one line) + hidden widget label (same full text for aria/JS)
    st.markdown(f'<div class="one-line-caption"><strong>{CA_LABEL}</strong></div>', unsafe_allow_html=True)
    st.text_input(
        CA_LABEL,
        value=st.session_state.get("ca_field", ""),
        key="ca_field",
        label_visibility="collapsed",
    )

    # 2) Geo (same grey as other inputs) — label stays as-is
    st.selectbox("Гео", geo_list, index=0 if geo_list else 0, key="geo_field")

    # 3) Type + Days on ONE row, aligned; keep full caption for days on one line
    c1, c2 = st.columns([1, 1], gap="small")
    with c1:
        st.selectbox("Тип площадки", ["Площадка", "Концерт", "Фестиваль"], index=0, key="venue_type_field")
    with c2:
        st.markdown(f'<div class="one-line-caption"><strong>{DAYS_LABEL}</strong></div>', unsafe_allow_html=True)
        st.number_input(
            DAYS_LABEL,
            min_value=0,
            step=1,
            value=0,
            key="venue_days_field",
            label_visibility="collapsed",
        )

    # 4) Visitors plan — label stays as-is
    st.number_input("План посетителей (в тысячах человек)", min_value=0.0, step=0.1, value=0.0, key="visitors_plan_field")

# ---------------- Robust Excel-like colors + fixed widths (JS by aria-label) ----------------
# We style actual aria-labeled controls, not wrappers.
components.html(
    f"""
    <script>
      const root = window.parent.document;

      function setWidth(el) {{
        if (!el) return;
        el.style.width = "{FILTER_W}px";
        el.style.maxWidth = "{FILTER_W}px";
      }}

      function styleControl(label, bg, fg, bold=false, readonly=false) {{
        const el = root.querySelector(`[aria-label="${{label}}"]`);
        if (!el) return;

        // Control itself
        setWidth(el);
        el.style.background = bg;
        el.style.color = fg;
        if (bold) el.style.fontWeight = "600";
        if (readonly) {{
          el.setAttribute("readonly", "true");
          el.style.cursor = "not-allowed";
        }}

        // Streamlit widget container
        const w = el.closest('.stTextInput, .stNumberInput, .stSelectbox');
        if (w) setWidth(w);

        // Selectbox: baseweb wrapper (visible control)
        const sel = el.closest('div[data-baseweb="select"]');
        if (sel) {{
          setWidth(sel);
          const controls = sel.querySelectorAll('div');
          controls.forEach(d => {{
            d.style.background = bg;
            d.style.color = fg;
          }});
        }}

        // Sometimes baseweb wraps deeper; try inner selects too
        if (w) {{
          const innerSelect = w.querySelector('div[data-baseweb="select"]');
          if (innerSelect) {{
            setWidth(innerSelect);
            const controls2 = innerSelect.querySelectorAll('div');
            controls2.forEach(d => {{
              d.style.background = bg;
              d.style.color = fg;
            }});
          }}
        }}
      }}

      // Apply styles (Excel semantics)
      styleControl("{CA_LABEL}", "{AUTO_BLUE}", "#ffffff", true, true);
      styleControl("Гео", "{USER_GREY}", "#111111", false, false);
      styleControl("Тип площадки", "{USER_GREY}", "#111111", false, false);
      styleControl("{DAYS_LABEL}", "{USER_GREY}", "#111111", false, false);
      styleControl("План посетителей (в тысячах человек)", "{USER_GREY}", "#111111", false, false);

      // Tighten spacing in the left block
      const widgets = root.querySelectorAll(".stTextInput, .stNumberInput, .stSelectbox");
      widgets.forEach(w => { w.style.marginBottom = "10px"; });
    </script>
    """,
    height=0,
)

st.markdown("---")

# ---------------- Table (unchanged for this step) ----------------
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
            cellStyle={"backgroundColor": AUTO_BLUE, "color": "white"},
        )

for col_name in df.columns:
    if col_name not in READ_ONLY_COLS:
        gb.configure_column(
            col_name,
            editable=True,
            cellStyle={"backgroundColor": USER_GREY, "color": "black"},
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

    st.session_state["ca_field"] = "CALC"

    st.success("Расчёт выполнен (в этом шаге: фильтры — подписи в 1 строку, цвета, ширина 250px, выравнивание).")

    AgGrid(
        edited,
        gridOptions=gb.build(),
        height=650,
        theme="streamlit",
        allow_unsafe_jscode=True,
    )
