
import os
import pandas as pd
import openpyxl
import streamlit as st

APP_TITLE = "Калькулятор оценки мероприятий"

# --- File path: keep everything in one folder (repo root)
DEFAULT_XLSX = "Калькулятор_оценки_мероприятий_по_городам.xlsx"

# --- Excel helpers

def load_workbook(path: str):
    # data_only=False: we need formulas as strings; calculation will be implemented in v2
    return openpyxl.load_workbook(path, data_only=False)


def get_cities(wb: openpyxl.Workbook) -> list[str]:
    ws = wb["ЦА по городам"]
    cities = []
    # expect city names in column A starting row 2
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        v = str(v).strip()
        if v:
            cities.append(v)
    # unique preserve order
    seen = set()
    out = []
    for c in cities:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def template_table_to_df(wb: openpyxl.Workbook, sheet_name: str = "TEMPLATE") -> pd.DataFrame:
    ws = wb[sheet_name]

    # Heuristic display range based on the template bbox we observed
    # Rows 27..134: the main calculator table block
    r1, r2 = 27, 134
    # Columns C..O (3..15)
    c1, c2 = 3, 15

    data = []
    for r in range(r1, r2 + 1):
        row = []
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            # Keep user-friendly values; for formulas keep as "" until calculated
            if isinstance(v, str) and v.startswith("="):
                row.append("")
            else:
                row.append(v)
        data.append(row)

    cols = [openpyxl.utils.get_column_letter(c) for c in range(c1, c2 + 1)]
    df = pd.DataFrame(data, columns=cols)
    # Add Excel row numbers for orientation
    df.insert(0, "ROW", list(range(r1, r2 + 1)))
    return df


def render_editor(df: pd.DataFrame) -> pd.DataFrame:
    # Columns that are typically user inputs in the Excel template.
    # We keep them editable; the rest read-only via disabling columns.
    # NOTE: streamlit data_editor disables by column, not by cell.
    editable_cols = {"D", "E", "F", "G", "H", "I"}  # channel/format/period/branding/manual OTS/manual reach
    disabled_cols = [c for c in df.columns if c != "ROW" and c not in editable_cols]

    st.caption("Серые поля ввода (как в Excel) доступны для редактирования. Голубые поля расчёта будут заполняться после нажатия ‘Рассчитать’ (в v1 расчёт пока не выполняется — это UI-скелет).")

    # Simple styling: inputs vs outputs
    def style_row(row):
        styles = {}
        for col in df.columns:
            if col == "ROW":
                styles[col] = "background-color: #F7F7F7;"
            elif col in editable_cols:
                styles[col] = "background-color: #E6E6E6;"  # grey
            else:
                styles[col] = "background-color: #D9EEF9;"  # light blue
        return pd.Series(styles)

    styled = df.style.apply(style_row, axis=1)

    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        disabled=disabled_cols,
        num_rows="fixed",
    )
    return edited


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    xlsx_path = DEFAULT_XLSX
    if not os.path.exists(xlsx_path):
        st.error(f"Не найден файл источника рядом с приложением: {xlsx_path}")
        st.stop()

    wb = load_workbook(xlsx_path)

    cities = get_cities(wb)
    if not cities:
        st.error("Не удалось прочитать список городов с листа 'ЦА по городам' (ожидается колонка A, начиная со строки 2).")
        st.stop()

    # --- Top controls
    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        city = st.selectbox("Город", cities)
    with col2:
        scenario_name = st.text_input("Название сценария", value="Сценарий 1")
    with col3:
        st.info("v1: первая версия интерфейса. Полный расчёт 1-в-1 как в Excel будет добавлен следующим шагом.")

    st.divider()

    # --- Parameters block (minimal)
    st.subheader("Параметры")
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        days = st.number_input("Дней мероприятия", min_value=0, value=0, step=1)
    with p2:
        period = st.number_input("Период размещения (дней)", min_value=0, value=0, step=1)
    with p3:
        visitors = st.number_input("План посетителей", min_value=0.0, value=0.0, step=1000.0)
    with p4:
        st.text_input("Выбранный город", value=city, disabled=True)

    st.divider()

    # --- Main table
    st.subheader("Таблица (как в Excel)")
    df = template_table_to_df(wb, "TEMPLATE")

    edited_df = render_editor(df)

    # --- Calculate
    if st.button("Рассчитать", type="primary"):
        # v1: calculation engine is not implemented; we just confirm inputs were captured.
        st.success("UI-версия v1: данные ввода сохранены. В следующей итерации сюда будет добавлен расчёт и заполнение голубых полей.")

        # Show a small preview of edited inputs for debugging
        st.subheader("Проверка ввода")
        show_cols = ["ROW", "D", "E", "F", "G", "H", "I"]
        show_cols = [c for c in show_cols if c in edited_df.columns]
        st.dataframe(edited_df[show_cols], use_container_width=True)


if __name__ == "__main__":
    main()
