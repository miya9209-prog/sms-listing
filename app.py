import io
import os
import re
import sqlite3
from datetime import datetime

import pandas as pd
import streamlit as st

APP_TITLE = "광고발송 고객리스트 생성"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
SAVED_DIR = os.path.join(DATA_DIR, "saved_lists")
DB_PATH = os.path.join(DATA_DIR, "saved_lists.db")

os.makedirs(SAVED_DIR, exist_ok=True)

st.set_page_config(page_title=APP_TITLE, layout="wide")


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS saved_lists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            list_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            row_count INTEGER DEFAULT 0,
            created_at TEXT NOT NULL,
            source_filename TEXT
        )
        """
    )
    conn.commit()
    conn.close()


init_db()


COLUMN_ALIASES = {
    "payment_amount": ["실결제금액"],
    "payment_rate": ["실결제율"],
    "order_unit": ["주문당단가", "1회주문당금액"],
    "order_date": ["주문일", "최종주문일"],
    "visit_date": ["접속일", "최종접속일"],
    "sms_opt_in": ["모바일 메시지 수신여부"],
    "bad_member": ["불량회원"],
}


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes, filename: str):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    return df, sheet_name



def find_column(df: pd.DataFrame, candidates):
    cols = {str(c).strip(): c for c in df.columns}
    for candidate in candidates:
        if candidate in cols:
            return cols[candidate]
    return None



def coerce_numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": None, "nan": None, "None": None})
    return pd.to_numeric(cleaned, errors="coerce")



def coerce_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")



def normalize_sms(value):
    if pd.isna(value):
        return False
    text = str(value).strip().upper()
    return text in {"T", "Y", "TRUE", "1", "동의", "수신"}



def normalize_bad_member(value):
    if pd.isna(value):
        return False
    text = str(value).strip().upper()
    return text in {"T", "Y", "TRUE", "1", "불량", "YES"}



def apply_filters(
    df: pd.DataFrame,
    payment_min,
    payment_max,
    rate_min,
    rate_max,
    unit_min,
    unit_max,
    order_from,
    order_to,
    visit_from,
    visit_to,
    sms_only,
    exclude_bad,
):
    filtered = df.copy()

    col_payment = find_column(filtered, COLUMN_ALIASES["payment_amount"])
    col_rate = find_column(filtered, COLUMN_ALIASES["payment_rate"])
    col_unit = find_column(filtered, COLUMN_ALIASES["order_unit"])
    col_order = find_column(filtered, COLUMN_ALIASES["order_date"])
    col_visit = find_column(filtered, COLUMN_ALIASES["visit_date"])
    col_sms = find_column(filtered, COLUMN_ALIASES["sms_opt_in"])
    col_bad = find_column(filtered, COLUMN_ALIASES["bad_member"])

    if col_payment:
        vals = coerce_numeric(filtered[col_payment])
        if payment_min is not None:
            filtered = filtered[vals >= payment_min]
            vals = coerce_numeric(filtered[col_payment])
        if payment_max is not None:
            filtered = filtered[vals <= payment_max]

    if col_rate:
        vals = coerce_numeric(filtered[col_rate])
        # 실결제율이 0~1로 저장된 경우 0~100 입력값과 맞춤
        if vals.dropna().max() is not None and len(vals.dropna()) > 0 and vals.dropna().max() <= 1.0:
            vals = vals * 100
        if rate_min is not None:
            filtered = filtered[vals >= rate_min]
            vals = coerce_numeric(filtered[col_rate])
            if vals.dropna().max() is not None and len(vals.dropna()) > 0 and vals.dropna().max() <= 1.0:
                vals = vals * 100
        if rate_max is not None:
            filtered = filtered[vals <= rate_max]

    if col_unit:
        vals = coerce_numeric(filtered[col_unit])
        if unit_min is not None:
            filtered = filtered[vals >= unit_min]
            vals = coerce_numeric(filtered[col_unit])
        if unit_max is not None:
            filtered = filtered[vals <= unit_max]

    if col_order:
        vals = coerce_datetime(filtered[col_order]).dt.date
        if order_from is not None:
            filtered = filtered[vals >= order_from]
            vals = coerce_datetime(filtered[col_order]).dt.date
        if order_to is not None:
            filtered = filtered[vals <= order_to]

    if col_visit:
        vals = coerce_datetime(filtered[col_visit]).dt.date
        if visit_from is not None:
            filtered = filtered[vals >= visit_from]
            vals = coerce_datetime(filtered[col_visit]).dt.date
        if visit_to is not None:
            filtered = filtered[vals <= visit_to]

    if sms_only and col_sms:
        filtered = filtered[filtered[col_sms].apply(normalize_sms)]

    if exclude_bad and col_bad:
        filtered = filtered[~filtered[col_bad].apply(normalize_bad_member)]

    return filtered



def get_sort_options(df: pd.DataFrame):
    options = {}
    mapping = {
        "실결제금액": COLUMN_ALIASES["payment_amount"],
        "실결제율": COLUMN_ALIASES["payment_rate"],
        "주문당단가": COLUMN_ALIASES["order_unit"],
        "주문일": COLUMN_ALIASES["order_date"],
        "접속일": COLUMN_ALIASES["visit_date"],
    }
    for label, candidates in mapping.items():
        col = find_column(df, candidates)
        if col:
            options[label] = col
    return options



def sort_dataframe(df: pd.DataFrame, sort_col: str, ascending: bool):
    if sort_col is None:
        return df
    sorted_df = df.copy()
    if sort_col in COLUMN_ALIASES["order_date"] + COLUMN_ALIASES["visit_date"]:
        sorted_df["__sort_temp__"] = coerce_datetime(sorted_df[sort_col])
    else:
        temp = coerce_numeric(sorted_df[sort_col])
        if sort_col in COLUMN_ALIASES["payment_rate"] and len(temp.dropna()) > 0 and temp.dropna().max() <= 1.0:
            temp = temp * 100
        sorted_df["__sort_temp__"] = temp
    sorted_df = sorted_df.sort_values(by="__sort_temp__", ascending=ascending, na_position="last")
    return sorted_df.drop(columns=["__sort_temp__"])



def dataframe_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "리스트") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        ws.freeze_panes = "A2"
    output.seek(0)
    return output.getvalue()



def sanitize_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name).strip()
    name = re.sub(r"\s+", "_", name)
    return name[:80] or "saved_list"



def save_list_to_db(list_name: str, df: pd.DataFrame, source_filename: str):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = sanitize_filename(list_name)
    filename = f"{safe_name}_{timestamp}.xlsx"
    file_path = os.path.join(SAVED_DIR, filename)
    xlsx_bytes = dataframe_to_xlsx_bytes(df)
    with open(file_path, "wb") as f:
        f.write(xlsx_bytes)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO saved_lists (list_name, file_path, row_count, created_at, source_filename) VALUES (?, ?, ?, ?, ?)",
        (list_name, file_path, len(df), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), source_filename),
    )
    conn.commit()
    conn.close()



def load_saved_lists():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(
        "SELECT id, list_name, file_path, row_count, created_at, source_filename FROM saved_lists ORDER BY id DESC",
        conn,
    )
    conn.close()
    return df



def delete_saved_list(list_id: int, file_path: str):
    if os.path.exists(file_path):
        os.remove(file_path)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DELETE FROM saved_lists WHERE id = ?", (list_id,))
    conn.commit()
    conn.close()


st.title(APP_TITLE)
st.caption("엑셀 파일을 업로드하고 조건별로 고객을 추출한 뒤, 정렬하여 xlsx로 다운로드하거나 이름을 붙여 저장할 수 있습니다.")

uploaded_file = st.file_uploader("고객 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file is not None:
    raw_bytes = uploaded_file.getvalue()
    df, sheet_name = load_excel(raw_bytes, uploaded_file.name)

    sort_options = get_sort_options(df)

    with st.expander("필터 조건", expanded=True):
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("**금액/비율 조건**")
            payment_min = st.number_input("실주문금액 시작", min_value=0, value=0, step=1000)
            payment_max_raw = st.number_input("실주문금액 끝", min_value=0, value=0, step=1000)
            rate_min = st.number_input("실결제율 시작(%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
            rate_max_raw = st.number_input("실결제율 끝(%)", min_value=0.0, max_value=100.0, value=100.0, step=1.0)
            unit_min = st.number_input("주문당단가 시작", min_value=0, value=0, step=1000)
            unit_max_raw = st.number_input("주문당단가 끝", min_value=0, value=0, step=1000)

        with c2:
            st.markdown("**기간/추가 조건**")
            today = datetime.now().date()
            order_from = st.date_input("주문일 시작", value=None)
            order_to = st.date_input("주문일 끝", value=today)
            visit_from = st.date_input("접속일 시작", value=None)
            visit_to = st.date_input("접속일 끝", value=today)
            sms_only = st.checkbox("문자수신동의 고객만", value=True)
            exclude_bad = st.checkbox("불량회원 제외", value=True)

    with st.expander("정렬 설정", expanded=True):
        c3, c4 = st.columns([2, 1])
        with c3:
            sort_label = st.selectbox("정렬 항목", options=["선택 안함"] + list(sort_options.keys()))
        with c4:
            sort_order = st.selectbox("정렬 방향", options=["내림차순", "오름차순"])

    payment_max = None if payment_max_raw == 0 else payment_max_raw
    rate_max = None if rate_max_raw == 100.0 else rate_max_raw
    unit_max = None if unit_max_raw == 0 else unit_max_raw
    payment_min_val = None if payment_min == 0 else payment_min
    rate_min_val = None if rate_min == 0 else rate_min
    unit_min_val = None if unit_min == 0 else unit_min

    filtered_df = apply_filters(
        df=df,
        payment_min=payment_min_val,
        payment_max=payment_max,
        rate_min=rate_min_val,
        rate_max=rate_max,
        unit_min=unit_min_val,
        unit_max=unit_max,
        order_from=order_from,
        order_to=order_to,
        visit_from=visit_from,
        visit_to=visit_to,
        sms_only=sms_only,
        exclude_bad=exclude_bad,
    )

    selected_sort_col = None if sort_label == "선택 안함" else sort_options.get(sort_label)
    result_df = sort_dataframe(filtered_df, selected_sort_col, ascending=(sort_order == "오름차순"))

    c5, c6, c7 = st.columns(3)
    c5.metric("원본 행 수", len(df))
    c6.metric("추출 행 수", len(result_df))
    c7.metric("제외 행 수", len(df) - len(result_df))

    st.dataframe(result_df, use_container_width=True, height=520)

    dl_name = f"광고발송_고객리스트_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "현재 리스트 xlsx 다운로드",
        data=dataframe_to_xlsx_bytes(result_df),
        file_name=dl_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()
    st.subheader("리스트 저장")
    save_name = st.text_input("저장할 리스트 이름", placeholder="예: 고액결제_최근접속_문자동의")
    if st.button("현재 리스트 저장", use_container_width=True):
        if not save_name.strip():
            st.warning("리스트 이름을 입력해 주세요.")
        else:
            save_list_to_db(save_name.strip(), result_df, uploaded_file.name)
            st.success(f"'{save_name.strip()}' 리스트를 저장했습니다.")
            st.rerun()

st.divider()
st.subheader("저장된 리스트 DB")
saved_df = load_saved_lists()

if saved_df.empty:
    st.info("저장된 리스트가 없습니다.")
else:
    st.dataframe(saved_df[["id", "list_name", "row_count", "created_at", "source_filename"]], use_container_width=True, height=260)
    saved_options = {
        f"[{row['id']}] {row['list_name']} ({row['row_count']}건, {row['created_at']})": row
        for _, row in saved_df.iterrows()
    }
    selected_saved_label = st.selectbox("불러올 저장 리스트 선택", options=list(saved_options.keys()))
    selected_row = saved_options[selected_saved_label]

    if os.path.exists(selected_row["file_path"]):
        preview_df = pd.read_excel(selected_row["file_path"])
        st.markdown(f"**미리보기: {selected_row['list_name']}**")
        st.dataframe(preview_df, use_container_width=True, height=360)

        with open(selected_row["file_path"], "rb") as f:
            st.download_button(
                "선택 리스트 xlsx 다운로드",
                data=f.read(),
                file_name=os.path.basename(selected_row["file_path"]),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if st.button("선택 리스트 삭제", use_container_width=True):
            delete_saved_list(int(selected_row["id"]), selected_row["file_path"])
            st.success("선택한 저장 리스트를 삭제했습니다.")
            st.rerun()
    else:
        st.warning("저장 파일을 찾을 수 없습니다. 목록에서 삭제 후 다시 저장해 주세요.")
