import io
import os
import sqlite3
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="광고발송 고객리스트 생성", layout="wide")

APP_TITLE = "광고발송 고객리스트 생성"
DATA_DIR = "data"
DB_PATH = os.path.join(DATA_DIR, "saved_lists.db")
EXPORT_DIR = os.path.join(DATA_DIR, "saved_lists")

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)


@st.cache_data(show_spinner=False)
def load_excel(uploaded_file):
    return pd.read_excel(uploaded_file)


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS saved_lists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            list_name TEXT UNIQUE NOT NULL,
            file_path TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    return conn


def normalize_currency(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9.-]", "", regex=True)
        .replace("", pd.NA)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_percent(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(r"[^0-9.-]", "", regex=True)
        .replace("", pd.NA)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.date


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="리스트")
    output.seek(0)
    return output.getvalue()


def save_list_to_db(name: str, df: pd.DataFrame):
    safe_name = "".join(ch for ch in name if ch not in '\\/:*?"<>|').strip()
    if not safe_name:
        raise ValueError("저장 이름을 입력해 주세요.")

    file_path = os.path.join(EXPORT_DIR, f"{safe_name}.xlsx")
    with open(file_path, "wb") as f:
        f.write(to_excel_bytes(df))

    conn = get_conn()
    try:
        conn.execute(
            """
            INSERT INTO saved_lists (list_name, file_path, row_count, created_at)
            VALUES (?, ?, ?, datetime('now', 'localtime'))
            ON CONFLICT(list_name) DO UPDATE SET
                file_path=excluded.file_path,
                row_count=excluded.row_count,
                created_at=excluded.created_at
            """,
            (safe_name, file_path, len(df)),
        )
        conn.commit()
    finally:
        conn.close()


def load_saved_lists() -> pd.DataFrame:
    conn = get_conn()
    try:
        return pd.read_sql_query(
            "SELECT id, list_name, file_path, row_count, created_at FROM saved_lists ORDER BY created_at DESC",
            conn,
        )
    finally:
        conn.close()


def delete_saved_list(list_name: str):
    conn = get_conn()
    try:
        row = conn.execute(
            "SELECT file_path FROM saved_lists WHERE list_name = ?", (list_name,)
        ).fetchone()
        if row and row[0] and os.path.exists(row[0]):
            os.remove(row[0])
        conn.execute("DELETE FROM saved_lists WHERE list_name = ?", (list_name,))
        conn.commit()
    finally:
        conn.close()


st.title(APP_TITLE)
st.caption("엑셀 업로드 → 조건 선택 → 리스트 추출 → xlsx 저장/다운로드")

uploaded_file = st.file_uploader("견본 또는 고객 xlsx 업로드", type=["xlsx"])

if uploaded_file is not None:
    try:
        raw_df = load_excel(uploaded_file)
    except Exception as e:
        st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {e}")
        st.stop()

    st.success(f"업로드 완료: {uploaded_file.name} / {len(raw_df):,}행")

    amount_col = find_column(raw_df, ["실주문금액", "실결제금액"])
    rate_col = find_column(raw_df, ["실결제율"])
    unit_col = find_column(raw_df, ["주문당단가", "1회주문당금액"])
    order_col = find_column(raw_df, ["주문일", "최종주문일"])
    access_col = find_column(raw_df, ["접속일", "최종접속일", "최종접속 일"])
    sms_col = find_column(raw_df, ["문자수신동의", "SMS수신동의", "sms수신동의"])
    bad_col = find_column(raw_df, ["회원등급", "회원상태", "고객상태"])

    with st.expander("필터 조건", expanded=True):
        c1, c2 = st.columns(2)

        with c1:
            st.markdown("#### 금액/비율 조건")

            a1, a2 = st.columns([1, 1])
            with a1:
                amount_min = st.number_input("실주문금액 시작", min_value=0, value=0, step=1000)
            with a2:
                amount_max = st.number_input("실주문금액 끝", min_value=0, value=0, step=1000)

            b1, b2 = st.columns([1, 1])
            with b1:
                rate_min = st.number_input("실결제율 시작(%)", min_value=0.0, max_value=100.0, value=0.0, step=0.1)
            with b2:
                rate_max = st.number_input("실결제율 끝(%)", min_value=0.0, max_value=100.0, value=100.0, step=0.1)

            d1, d2 = st.columns([1, 1])
            with d1:
                unit_min = st.number_input("주문당단가 시작", min_value=0, value=0, step=1000)
            with d2:
                unit_max = st.number_input("주문당단가 끝", min_value=0, value=0, step=1000)

        with c2:
            st.markdown("#### 기간/추가 조건")

            o1, o2 = st.columns([1, 1])
            with o1:
                order_start = st.date_input("주문일 시작", value=None)
            with o2:
                order_end = st.date_input("주문일 끝", value=date.today())

            v1, v2 = st.columns([1, 1])
            with v1:
                access_start = st.date_input("접속일 시작", value=None)
            with v2:
                access_end = st.date_input("접속일 끝", value=date.today())

            sms_only = st.checkbox("문자수신동의 고객만", value=True)
            exclude_bad = st.checkbox("불량회원 제외", value=True)

    sort_options = [x for x in [amount_col, rate_col, unit_col, order_col, access_col] if x]
    default_sort = sort_options[0] if sort_options else None
    s1, s2 = st.columns([2, 1])
    with s1:
        sort_col = st.selectbox("정렬 기준", options=sort_options if sort_options else ["정렬 불가"], index=0)
    with s2:
        sort_ascending = st.selectbox("정렬 방식", options=["내림차순", "오름차순"], index=0)

    if st.button("리스트 생성", use_container_width=True):
        result = raw_df.copy()

        if amount_col:
            vals = normalize_currency(result[amount_col])
            if amount_max > 0:
                result = result[(vals >= amount_min) & (vals <= amount_max)]
            else:
                result = result[vals >= amount_min]

        if rate_col:
            vals = normalize_percent(result[rate_col])
            result = result[(vals >= rate_min) & (vals <= rate_max)]

        if unit_col:
            vals = normalize_currency(result[unit_col])
            if unit_max > 0:
                result = result[(vals >= unit_min) & (vals <= unit_max)]
            else:
                result = result[vals >= unit_min]

        if order_col:
            vals = normalize_date(result[order_col])
            if order_start:
                result = result[vals >= order_start]
                vals = normalize_date(result[order_col])
            if order_end:
                result = result[vals <= order_end]

        if access_col:
            vals = normalize_date(result[access_col])
            if access_start:
                result = result[vals >= access_start]
                vals = normalize_date(result[access_col])
            if access_end:
                result = result[vals <= access_end]

        if sms_only and sms_col:
            sms_vals = result[sms_col].astype(str).str.strip().str.upper()
            result = result[sms_vals.isin(["Y", "YES", "TRUE", "1", "동의"])]

        if exclude_bad and bad_col:
            bad_vals = result[bad_col].astype(str).str.strip()
            result = result[~bad_vals.str.contains("불량", na=False)]

        if sort_options and sort_col != "정렬 불가":
            if sort_col in [amount_col, unit_col]:
                temp_sort = normalize_currency(result[sort_col])
            elif sort_col == rate_col:
                temp_sort = normalize_percent(result[sort_col])
            elif sort_col in [order_col, access_col]:
                temp_sort = pd.to_datetime(result[sort_col], errors="coerce")
            else:
                temp_sort = result[sort_col]

            result = result.assign(_sort_key=temp_sort).sort_values(
                by="_sort_key",
                ascending=(sort_ascending == "오름차순"),
                na_position="last",
            ).drop(columns=["_sort_key"])

        st.session_state["result_df"] = result
        st.success(f"조건에 맞는 고객 {len(result):,}명")

    if "result_df" in st.session_state:
        result = st.session_state["result_df"]
        st.dataframe(result, use_container_width=True, height=420)

        excel_bytes = to_excel_bytes(result)
        st.download_button(
            "xlsx 다운로드",
            data=excel_bytes,
            file_name="광고발송_고객리스트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        save_name = st.text_input("DB 저장 이름", placeholder="예: 4월 재구매 유도 고객")
        if st.button("현재 리스트 저장", use_container_width=True):
            try:
                save_list_to_db(save_name, result)
                st.success("리스트를 저장했습니다.")
            except Exception as e:
                st.error(str(e))

st.divider()
st.subheader("저장된 리스트")
saved_df = load_saved_lists()

if saved_df.empty:
    st.info("아직 저장된 리스트가 없습니다.")
else:
    selected_name = st.selectbox("저장 리스트 선택", options=saved_df["list_name"].tolist())
    selected_row = saved_df[saved_df["list_name"] == selected_name].iloc[0]

    info1, info2 = st.columns(2)
    with info1:
        st.caption(f"저장일시: {selected_row['created_at']}")
    with info2:
        st.caption(f"행 수: {int(selected_row['row_count']):,}행")

    action1, action2 = st.columns(2)
    with action1:
        if st.button("불러오기", use_container_width=True):
            try:
                loaded_df = pd.read_excel(selected_row["file_path"])
                st.session_state["loaded_saved_df"] = loaded_df
                st.success("저장 리스트를 불러왔습니다.")
            except Exception as e:
                st.error(f"불러오기 실패: {e}")
    with action2:
        if st.button("선택 리스트 삭제", use_container_width=True):
            delete_saved_list(selected_name)
            st.success("삭제되었습니다. 새로고침하면 목록이 갱신됩니다.")

    if "loaded_saved_df" in st.session_state:
        loaded_df = st.session_state["loaded_saved_df"]
        st.dataframe(loaded_df, use_container_width=True, height=320)
        st.download_button(
            "불러온 리스트 xlsx 다운로드",
            data=to_excel_bytes(loaded_df),
            file_name=f"{selected_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
