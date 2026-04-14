import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(layout="wide")
st.title("광고발송 고객리스트 생성")

uploaded_file = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("필터 조건")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 금액/비율")

        c1, c2 = st.columns(2)
        with c1:
            amount_min = st.number_input("실주문금액 시작", 0)
        with c2:
            amount_max = st.number_input("실주문금액 끝", 999999999)

        c1, c2 = st.columns(2)
        with c1:
            rate_min = st.number_input("실결제율 시작(%)", 0.0)
        with c2:
            rate_max = st.number_input("실결제율 끝(%)", 100.0)

        c1, c2 = st.columns(2)
        with c1:
            unit_min = st.number_input("주문당단가 시작", 0)
        with c2:
            unit_max = st.number_input("주문당단가 끝", 999999999)

    with col2:
        st.markdown("### 기간")

        c1, c2 = st.columns(2)
        with c1:
            order_start = st.date_input("주문일 시작", None)
        with c2:
            order_end = st.date_input("주문일 끝", datetime.today())

        c1, c2 = st.columns(2)
        with c1:
            visit_start = st.date_input("접속일 시작", None)
        with c2:
            visit_end = st.date_input("접속일 끝", datetime.today())

        sms_only = st.checkbox("문자수신동의 고객만")
        exclude_bad = st.checkbox("불량회원 제외")

    if st.button("리스트 생성"):

        result = df.copy()

        # 🔥 실주문금액
        if "실주문금액" in result.columns:
            result = result[
                (result["실주문금액"] >= amount_min) &
                (result["실주문금액"] <= amount_max)
            ]

        # 🔥 실결제율 (0~1 → 0~100 자동 보정)
        if "실결제율" in result.columns:
            if result["실결제율"].max() <= 1:
                result["실결제율"] = result["실결제율"] * 100

            result = result[
                (result["실결제율"] >= rate_min) &
                (result["실결제율"] <= rate_max)
            ]

        # 🔥 주문당단가
        if "1회주문당금액" in result.columns:
            result = result[
                (result["1회주문당금액"] >= unit_min) &
                (result["1회주문당금액"] <= unit_max)
            ]

        # 🔥 주문일
        if "주문일" in result.columns and order_start:
            result["주문일"] = pd.to_datetime(result["주문일"]).dt.date
            result = result[
                (result["주문일"] >= order_start) &
                (result["주문일"] <= order_end)
            ]

        # 🔥 접속일
        if "접속일" in result.columns and visit_start:
            result["접속일"] = pd.to_datetime(result["접속일"]).dt.date
            result = result[
                (result["접속일"] >= visit_start) &
                (result["접속일"] <= visit_end)
            ]

        # 🔥 문자수신동의
        if sms_only and "모바일 메시지 수신여부" in result.columns:
            result = result[result["모바일 메시지 수신여부"] == "Y"]

        # 🔥 불량회원 제외
        if exclude_bad and "불량회원" in result.columns:
            result = result[result["불량회원"] != "Y"]

        st.success(f"{len(result)}명 추출됨")

        if len(result) == 0:
            st.warning("조건이 너무 좁습니다. 일부 조건을 줄여보세요.")

        st.dataframe(result)

        file_name = "result.xlsx"
        result.to_excel(file_name, index=False)

        with open(file_name, "rb") as f:
            st.download_button("xlsx 다운로드", f, file_name)
