import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

APP_TITLE = "City Events Calculator"
WORKBOOK_PATH = "Calculator.xlsm"  # keep the xlsm рядом с app.py

# UI work range (per Step A)
R1, C1 = 20, 3   # C20
R2, C2 = 146, 15 # O146

# Blocks (for better UX)
PARAM_R1, PARAM_C1, PARAM_R2, PARAM_C2 = 20, 3, 26, 5    # C20:E26
MAIN_R1, MAIN_C1, MAIN_R2, MAIN_C2 = 27, 3, 134, 15     # C27:O134
BOT_R1,  BOT_C1,  BOT_R2,  BOT_C2  = 135, 3, 146, 15    # C135:O146

# Known fills (ARGB). We use these first; otherwise we fall back to "non-white".
KNOWN_GREYS = {"FFD9D9D9", "FFBFBFBF", "FFE7E6E6", "FFDDDDDD"}
KNOWN_BLUES = {"FFBDD7EE", "FF9DC3E6", "FFB4C6E7", "FFB7DEE8"}


# ------------------------------------------------------------
# Normalization (Step A/B)
# ------------------------------------------------------------

def normalize_key(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u00A0", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("«", '"').replace("»", '"').replace("“", '"').replace("”", '"')
    s = s.replace("–", "-").replace("—", "-")
    return s.upper()


# ------------------------------------------------------------
# Excel helpers
# ------------------------------------------------------------

def a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def safe_cell_value(cell) -> Optional[object]:
    """Return value without exposing formulas to UI."""
    v = cell.value
    if isinstance(v, str) and v.startswith("="):
        return None
    return v


def cell_fill_argb(cell) -> str:
    try:
        f = cell.fill
        if f is None or f.patternType is None:
            return ""
        c = f.fgColor
        if c is None:
            return ""
        if c.type == "rgb" and c.rgb:
            return c.rgb
        if c.type == "indexed" and c.indexed is not None:
            # indexed colors: we don’t resolve palette here
            return ""
    except Exception:
        return ""
    return ""


def is_grey_input(cell) -> bool:
    argb = cell_fill_argb(cell)
    if argb in KNOWN_GREYS:
        return True
    # Fallback: treat any non-white solid fill as potential "input",
    # but we will still validate editable cells by policy.
    if argb and argb not in KNOWN_BLUES and argb != "FFFFFFFF":
        # In this template grey inputs are the only solid non-blue fills.
        return True
    return False


def is_blue_auto(cell) -> bool:
    argb = cell_fill_argb(cell)
    return argb in KNOWN_BLUES


@dataclass
class ValidationList:
    options: List[str]


def extract_city_list(ws_lists) -> List[str]:
    # Column A, from row 2 until blank
    cities: List[str] = []
    r = 2
    while True:
        v = ws_lists[f"A{r}"].value
        if v is None or str(v).strip() == "":
            break
        cities.append(str(v).strip())
        r += 1
    return cities


def extract_type_ploshadki(ws_lists) -> List[str]:
    # W2:W4
    opts = []
    for r in range(2, 20):
        v = ws_lists[f"W{r}"].value
        if v is None or str(v).strip() == "":
            break
        opts.append(str(v).strip())
    return opts


def extract_union_formats(ws_lists) -> List[str]:
    # Union of all format columns mentioned in Step B map.
    cols = ["B", "C", "D", "E", "F", "H", "I", "K", "L", "M", "N", "O", "P", "Q"]
    opts = set()
    for col in cols:
        for r in range(2, 300):
            v = ws_lists[f"{col}{r}"].value
            if v is None or str(v).strip() == "":
                # stop if we reached long empty tail for this column
                if r > 30:
                    break
                continue
            opts.add(str(v).strip())
    return sorted(opts)


def sheet_candidates(wb) -> List[str]:
    # Prefer real city sheets if present; otherwise TEMPLATE.
    names = list(wb.sheetnames)
    if "TEMPLATE" in names:
        # show TEMPLATE as the only layout sheet
        return ["TEMPLATE"]
    return names


# ------------------------------------------------------------
# Rendering helpers
# ------------------------------------------------------------

def block_to_dataframe(ws_values, ws_style, r1: int, c1: int, r2: int, c2: int) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return (values_df, grey_mask_df, blue_mask_df)."""
    values = []
    grey_mask = []
    blue_mask = []
    for r in range(r1, r2 + 1):
        row_vals = []
        row_grey = []
        row_blue = []
        for c in range(c1, c2 + 1):
            cell_v = ws_values.cell(r, c)
            cell_s = ws_style.cell(r, c)
            row_vals.append(safe_cell_value(cell_v))
            row_grey.append(is_grey_input(cell_s))
            row_blue.append(is_blue_auto(cell_s))
        values.append(row_vals)
        grey_mask.append(row_grey)
        blue_mask.append(row_blue)

    cols = [get_column_letter(c) for c in range(c1, c2 + 1)]
    idx = list(range(r1, r2 + 1))
    vdf = pd.DataFrame(values, columns=cols, index=idx)
    gdf = pd.DataFrame(grey_mask, columns=cols, index=idx)
    bdf = pd.DataFrame(blue_mask, columns=cols, index=idx)
    return vdf, gdf, bdf


def style_block(df: pd.DataFrame, grey_mask: pd.DataFrame, blue_mask: pd.DataFrame):
    def _apply(_):
        out = pd.DataFrame("", index=df.index, columns=df.columns)
        out[grey_mask] = "background-color: #3a3a3a;"  # grey inputs
        out[blue_mask] = "background-color: #1f3a5f;"  # blue autos
        return out

    return df.style.apply(_apply, axis=None)


def coerce_numeric(s):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return s
    try:
        txt = str(s).strip().replace(" ", "").replace(",", ".")
        return float(txt)
    except Exception:
        return None


# ------------------------------------------------------------
# App
# ------------------------------------------------------------

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Excel-backed UI scaffold. Formulas are not recalculated in Python yet; values are taken from Excel cached results.")

    # Load two views:
    # - values: data_only=True (cached results)
    # - style:  data_only=False (formulas, fills, merges)
    wb_values = load_workbook(WORKBOOK_PATH, data_only=True, keep_vba=True)
    wb_style = load_workbook(WORKBOOK_PATH, data_only=False, keep_vba=True)

    if "Списки" not in wb_style.sheetnames:
        st.error("Не найден лист 'Списки' в книге.")
        return

    ws_lists = wb_style["Списки"]
    cities = extract_city_list(ws_lists)
    type_pl = extract_type_ploshadki(ws_lists)
    union_formats = extract_union_formats(ws_lists)

    # City selection (per your rule: pick city from list, layout is TEMPLATE)
    city = st.selectbox("Город (ГЕО)", options=cities, index=0)

    # We always render TEMPLATE (it is the canonical layout)
    sheet_name = "TEMPLATE" if "TEMPLATE" in wb_style.sheetnames else wb_style.sheetnames[0]
    ws_v = wb_values[sheet_name]
    ws_s = wb_style[sheet_name]

    # Parameters block
    st.subheader("Параметры")

    # Map: show the same labels as Excel; edit only grey inputs.
    # We take existing cached values; if empty — user can fill.
    # Inputs per Step A: E22:E26
    # Also keep E23 dropdown options.

    # Current cached values
    cur_e22 = ws_v["E22"].value or city
    cur_e23 = ws_v["E23"].value
    cur_e24 = ws_v["E24"].value
    cur_e25 = ws_v["E25"].value
    cur_e26 = ws_v["E26"].value

    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        geo_val = st.selectbox("Гео (E22)", options=cities, index=max(0, cities.index(city) if city in cities else 0))
        type_pl_val = st.selectbox("Тип площадки (E23)", options=type_pl if type_pl else ["Площадка"], index=0 if not cur_e23 else (type_pl.index(cur_e23) if cur_e23 in type_pl else 0))
    with col2:
        days_val = st.number_input("Кол-во дней на фестивале/площадке (E24)", value=float(coerce_numeric(cur_e24) or 0), step=1.0)
        period_val = st.number_input("Общий период размещения (E25)", value=float(coerce_numeric(cur_e25) or 0), step=1.0)
    with col3:
        visitors_val = st.number_input("План посетителей (E26), тыс.", value=float(coerce_numeric(cur_e26) or 0), step=1.0)
        st.info("ЦА (E21) и расчётные поля заполняются формулами Excel. Если вы видите пустые значения, откройте файл в Excel, нажмите пересчёт и сохраните — кэш значений появится.")

    # Main table
    st.subheader("Медиа факторы")

    vdf_main, gdf_main, bdf_main = block_to_dataframe(ws_v, ws_s, MAIN_R1, MAIN_C1, MAIN_R2, MAIN_C2)

    # Replace formulas with cached values; if missing, keep None.
    # For better UX: show empty string instead of None.
    display_main = vdf_main.copy()
    display_main = display_main.fillna("")

    # Provide editor, but restrict editing only to grey cells.
    edited = st.data_editor(
        display_main,
        use_container_width=True,
        height=620,
        column_config={
            # Column E (formats) — dropdown with union list (approximation)
            "E": st.column_config.SelectboxColumn("E", options=union_formats, required=False) if union_formats else None,
        },
        disabled=[c for c in display_main.columns if c not in display_main.columns],
        key="main_editor",
    )

    st.subheader("Итоги / нижний блок")
    vdf_bot, gdf_bot, bdf_bot = block_to_dataframe(ws_v, ws_s, BOT_R1, BOT_C1, BOT_R2, BOT_C2)
    display_bot = vdf_bot.fillna("")
    st.dataframe(style_block(display_bot, gdf_bot, bdf_bot), use_container_width=True, height=360)

    # Save
    st.divider()
    st.subheader("Сохранение")

    if st.button("Сохранить изменения в новый .xlsm"):
        # Write changes into a fresh workbook (style workbook, to preserve VBA/fills)
        wb_out = load_workbook(WORKBOOK_PATH, data_only=False, keep_vba=True)
        ws_out = wb_out[sheet_name]

        # Apply parameter inputs
        ws_out["E22"].value = geo_val
        ws_out["E23"].value = type_pl_val
        ws_out["E24"].value = int(days_val) if days_val.is_integer() else float(days_val)
        ws_out["E25"].value = int(period_val) if period_val.is_integer() else float(period_val)
        ws_out["E26"].value = int(visitors_val) if visitors_val.is_integer() else float(visitors_val)

        # Apply main table edits ONLY where cell is grey-input in Excel
        for r in range(MAIN_R1, MAIN_R2 + 1):
            for c in range(MAIN_C1, MAIN_C2 + 1):
                addr = a1(r, c)
                cell_s = ws_s[addr]
                if not is_grey_input(cell_s):
                    continue
                col_letter = get_column_letter(c)
                val = edited.loc[r, col_letter]
                # empty string -> clear
                if val == "":
                    ws_out[addr].value = None
                else:
                    ws_out[addr].value = val

        out = io.BytesIO()
        wb_out.save(out)
        out.seek(0)
        st.download_button(
            label="Скачать обновлённый файл (.xlsm)",
            data=out,
            file_name="Calculator_updated.xlsm",
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )


if __name__ == "__main__":
    main()
