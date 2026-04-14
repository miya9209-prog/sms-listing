import streamlit as st
import pandas as pd
from datetime import datetime
import sqlite3
import os

st.set_page_config(layout="wide")
st.title("광고발송 고객리스트 생성")

# DB 세팅
os.makedirs("data", exist_ok=True)
conn = sqlite3.connect("data/saved.db", check_same_thread=False)
c = conn.cursor()
c.execute("""CREATE TABLE IF NOT EXISTS lists
             (name TEXT, file BLOB)""")

# 업로드
uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("필터 조건")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 금액/비율")
        a_min = st.number_input("실주문금액 시작", 0)
        a_max = st.number_input("실주문금액 끝", 999999999)

        rate_min = st.number_input("실결제율 시작", 0.0)
        rate_max = st.number_input("실결제율 끝", 100.0)

        unit_min = st.number_input("주문당단가 시작", 0)
        unit_max = st.number_input("주문당단가 끝", 999999999)

    with col2:
        st.markdown("### 기간")
        order_start = st.date_input("주문일 시작", None)
        order_end = st.date_input("주문일 끝", datetime.today())

        visit_start = st.date_input("접속일 시작", None)
        visit_end = st.date_input("접속일 끝", datetime.today())

        sms_only = st.checkbox("문자수신동의 고객만")
        exclude_bad = st.checkbox("불량회원 제외")

    if st.button("리스트 생성"):

        result = df.copy()

        # 금액 필터
        if "실주문금액" in result.columns:
            result = result[(result["실주문금액"] >= a_min) & (result["실주문금액"] <= a_max)]

        if "실결제율" in result.columns:
            result = result[(result["실결제율"] >= rate_min) & (result["실결제율"] <= rate_max)]

        if "1회주문당금액" in result.columns:
            result = result[(result["1회주문당금액"] >= unit_min) & (result["1회주문당금액"] <= unit_max)]

        # 날짜 필터
        if "주문일" in result.columns and order_start:
            result["주문일"] = pd.to_datetime(result["주문일"]).dt.date
            result = result[result["주문일"] >= order_start]
            result = result[result["주문일"] <= order_end]

        if "접속일" in result.columns and visit_start:
            result["접속일"] = pd.to_datetime(result["접속일"]).dt.date
            result = result[result["접속일"] >= visit_start]
            result = result[result["접속일"] <= visit_end]

        # 체크 필터
        if sms_only and "문자수신동의" in result.columns:
            result = result[result["문자수신동의"] == "Y"]

        if exclude_bad and "회원등급" in result.columns:
            result = result[result["회원등급"] != "불량"]

        st.success(f"{len(result)}명 추출됨")
        st.dataframe(result)

        # 다운로드
        file_name = "result.xlsx"
        result.to_excel(file_name, index=False)

        with open(file_name, "rb") as f:
            st.download_button("xlsx 다운로드", f, file_name)

        # 저장
        name = st.text_input("DB 저장 이름")
        if st.button("저장"):
            result.to_excel("temp.xlsx", index=False)
            with open("temp.xlsx", "rb") as f:
                c.execute("INSERT INTO lists VALUES (?, ?)", (name, f.read()))
                conn.commit()
                st.success("저장 완료")

# 불러오기
st.divider()
st.subheader("저장 리스트")

rows = c.execute("SELECT name FROM lists").fetchall()
names = [r[0] for r in rows]

selected = st.selectbox("불러오기", names)

if st.button("로드"):
    row = c.execute("SELECT file FROM lists WHERE name=?", (selected,)).fetchone()
    with open("load.xlsx", "wb") as f:
        f.write(row[0])
    df_load = pd.read_excel("load.xlsx")
    st.dataframe(df_load)
