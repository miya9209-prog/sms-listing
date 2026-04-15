import io
import os
import re
import sqlite3
from typing import Optional

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


def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def normalize_currency(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .replace("", pd.NA)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def normalize_percent(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .replace("", pd.NA)
    )
    numeric = pd.to_numeric(cleaned, errors="coerce")
    valid = numeric.dropna()

    if valid.empty:
        return numeric

    ratio_small = ((valid >= 0) & (valid <= 1.5)).mean()
    median_val = valid.median()

    # 샘플 데이터처럼 대부분이 0~1 범위(예: 0.77 = 77%)일 때 자동 변환
    if ratio_small >= 0.9 and median_val <= 1.5:
        numeric = numeric * 100

    return numeric


def normalize_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.date


def parse_optional_number(text: str) -> Optional[float]:
    if text is None:
        return None
    s = str(text).strip()
    if s == "":
        return None

    s = re.sub(r"[^0-9.\-]", "", s)
    if s in {"", "-", ".", "-."}:
        return None

    return float(s)


def sms_true_mask(series: pd.Series) -> pd.Series:
    normalized = series.astype(str).str.strip().str.upper()
    return normalized.isin(["Y", "YES", "T", "TRUE", "1", "동의"])


def bad_member_mask(series: pd.Series) -> pd.Series:
    normalized = series.astype(str).str.strip().str.upper()
    return normalized.isin(["Y", "YES", "T", "TRUE", "1", "불량"])


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
            """
            SELECT id, list_name, file_path, row_count, created_at
            FROM saved_lists
            ORDER BY created_at DESC
            """,
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
    raw_df = load_excel(uploaded_file)
    st.success(f"업로드 완료: {uploaded_file.name} / {len(raw_df):,}행")

    amount_col = find_column(raw_df, ["실주문금액", "실결제금액"])
    rate_col = find_column(raw_df, ["실결제율"])
    unit_col = find_column(raw_df, ["주문당단가", "1회주문당금액"])
    order_col = find_column(raw_df, ["주문일", "최종주문일"])
    visit_col = find_column(raw_df, ["접속일", "최종접속일"])
    sms_col = find_column(raw_df, ["문자수신동의", "모바일 메시지 수신여부", "SMS수신동의"])
    bad_col = find_column(raw_df, ["불량회원", "회원상태"])

    with st.expander("필터 조건", expanded=True):
        left, right = st.columns(2)

        with left:
            st.markdown("### 금액/비율")

            a1, a2 = st.columns(2)
            amount_min_text = a1.text_input("실주문금액 시작", value="", placeholder="예: 100000")
            amount_max_text = a2.text_input("실주문금액 끝", value="", placeholder="예: 299999")

            r1, r2 = st.columns(2)
            rate_min_text = r1.text_input("실결제율 시작(%)", value="", placeholder="예: 70")
            rate_max_text = r2.text_input("실결제율 끝(%)", value="", placeholder="예: 100")

            u1, u2 = st.columns(2)
            unit_min_text = u1.text_input("주문당단가 시작", value="", placeholder="예: 0")
            unit_max_text = u2.text_input("주문당단가 끝", value="", placeholder="예: 150000")

        with right:
            st.markdown("### 기간")

            o1, o2 = st.columns(2)
            order_start = o1.date_input("주문일 시작", value=None, format="YYYY/MM/DD")
            order_end = o2.date_input("주문일 끝", value=None, format="YYYY/MM/DD")

            v1, v2 = st.columns(2)
            visit_start = v1.date_input("접속일 시작", value=None, format="YYYY/MM/DD")
            visit_end = v2.date_input("접속일 끝", value=None, format="YYYY/MM/DD")

            sms_only = st.checkbox("문자수신동의 고객만")
            exclude_bad = st.checkbox("불량회원 제외")

    sort_left, sort_right = st.columns([2, 1])
    sort_col = sort_left.selectbox(
        "정렬 기준",
        ["정렬 안함", "실주문금액", "실결제율", "주문당단가", "주문일", "접속일"],
        index=0,
    )
    sort_order = sort_right.selectbox("정렬 방식", ["내림차순", "오름차순"], index=0)

    if st.button("리스트 생성", use_container_width=True):
        try:
            amount_min = parse_optional_number(amount_min_text)
            amount_max = parse_optional_number(amount_max_text)
            rate_min = parse_optional_number(rate_min_text)
            rate_max = parse_optional_number(rate_max_text)
            unit_min = parse_optional_number(unit_min_text)
            unit_max = parse_optional_number(unit_max_text)
        except Exception:
            st.error("숫자 입력 형식을 확인해 주세요.")
            st.stop()

        result = raw_df.copy()

        if amount_col:
            amount_series = normalize_currency(result[amount_col])
            if amount_min is not None:
                result = result.loc[amount_series >= amount_min]
                amount_series = amount_series.loc[result.index]
            if amount_max is not None:
                result = result.loc[amount_series <= amount_max]

        if rate_col:
            rate_series = normalize_percent(result[rate_col])
            if rate_min is not None:
                result = result.loc[rate_series >= rate_min]
                rate_series = rate_series.loc[result.index]
            if rate_max is not None:
                result = result.loc[rate_series <= rate_max]

        if unit_col:
            unit_series = normalize_currency(result[unit_col])
            if unit_min is not None:
                result = result.loc[unit_series >= unit_min]
                unit_series = unit_series.loc[result.index]
            if unit_max is not None:
                result = result.loc[unit_series <= unit_max]

        if order_col:
            order_series = normalize_date(result[order_col])
            if order_start is not None:
                result = result.loc[order_series >= order_start]
                order_series = order_series.loc[result.index]
            if order_end is not None:
                result = result.loc[order_series <= order_end]

        if visit_col:
            visit_series = normalize_date(result[visit_col])
            if visit_start is not None:
                result = result.loc[visit_series >= visit_start]
                visit_series = visit_series.loc[result.index]
            if visit_end is not None:
                result = result.loc[visit_series <= visit_end]

        if sms_only and sms_col:
            result = result.loc[sms_true_mask(result[sms_col])]

        if exclude_bad and bad_col:
            result = result.loc[~bad_member_mask(result[bad_col])]

        sort_map = {
            "실주문금액": amount_col,
            "실결제율": rate_col,
            "주문당단가": unit_col,
            "주문일": order_col,
            "접속일": visit_col,
        }
        target_sort_col = sort_map.get(sort_col)
        if sort_col != "정렬 안함" and target_sort_col:
            if sort_col == "실결제율":
                sort_key = normalize_percent(result[target_sort_col])
            elif sort_col in ["실주문금액", "주문당단가"]:
                sort_key = normalize_currency(result[target_sort_col])
            elif sort_col in ["주문일", "접속일"]:
                sort_key = pd.to_datetime(result[target_sort_col], errors="coerce")
            else:
                sort_key = result[target_sort_col]

            result = (
                result.assign(_sort_key=sort_key)
                .sort_values("_sort_key", ascending=(sort_order == "오름차순"))
                .drop(columns=["_sort_key"])
            )

        st.success(f"조건에 맞는 고객 {len(result):,}명")

        if len(result) == 0:
            st.info(
                "이번 수정으로 실결제율 소수값(예: 0.77)을 77%로 자동 보정합니다. "
                "그래도 0명이면 조건 자체가 좁은 경우입니다."
            )

        st.dataframe(result, use_container_width=True, height=420)

        excel_bytes = to_excel_bytes(result)
        st.download_button(
            "xlsx 다운로드",
            data=excel_bytes,
            file_name="광고발송_고객리스트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown("### 리스트 저장")
        save_name = st.text_input("DB로 저장할 리스트 이름", value="", placeholder="예: 4월_문자발송_1차")
        if st.button("현재 결과 저장", use_container_width=True):
            try:
                save_list_to_db(save_name, result)
                st.success(f"'{save_name}' 저장 완료")
            except Exception as e:
                st.error(f"저장 중 오류가 발생했습니다: {e}")

st.divider()
st.subheader("저장된 리스트")

saved_df = load_saved_lists()
if saved_df.empty:
    st.caption("저장된 리스트가 없습니다.")
else:
    selected_name = st.selectbox("저장 리스트 선택", saved_df["list_name"].tolist())
    sel_row = saved_df[saved_df["list_name"] == selected_name].iloc[0]

    c1, c2, c3 = st.columns([1, 1, 1])
    c1.metric("행 수", f"{int(sel_row['row_count']):,}")
    c2.metric("저장 이름", str(sel_row["list_name"]))
    c3.metric("저장 시각", str(sel_row["created_at"]))

    file_path = sel_row["file_path"]
    if os.path.exists(file_path):
        preview_df = pd.read_excel(file_path)
        st.dataframe(preview_df, use_container_width=True, height=320)
        with open(file_path, "rb") as f:
            st.download_button(
                "저장 리스트 다운로드",
                data=f.read(),
                file_name=os.path.basename(file_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if st.button("선택 리스트 삭제", use_container_width=True):
        delete_saved_list(selected_name)
        st.success("삭제 완료")
        st.rerun()
