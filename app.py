# app.py
# -*- coding: utf-8 -*-
import streamlit as st
from tax_invoice_app import run as run_tax  # ← 분리한 파일에서 run() 가져오기
from misc_app import run as run_misc        # ← 기타기능 모듈

def go(page: str):
    """페이지 상태 변경 + 즉시 리렌더링"""
    st.session_state["page"] = page
    st.rerun()


def render_main_menu():
    st.title("📊 재무회계팀 자동화 작업 메뉴")
    st.write("원하는 작업을 선택하세요.")

    # ----- 버튼 공통 스타일: 너비 150px 고정 -----
    st.markdown(
        """
        <style>
            .small-button button {
                width: 150px !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    # ---------------------- 결산 작업 + 기타기능 ----------------------
    with col1:
        st.subheader("📘 결산 작업 📘")

        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        st.button("재무제표 생성", disabled=True)
        st.button("회계단위별 원장파일 통합", disabled=True)
        st.button("재무제표 vs 부속명세서 검증", disabled=True)

        st.markdown("---")   # 구분선

        st.subheader("🛠️ 기타기능 🛠️")
        if st.button("자금이체 적요 자동생성"):
            go("misc")
        st.markdown("</div>", unsafe_allow_html=True)

    # ---------------------- 검증 / 대조 ----------------------
    with col2:
        st.subheader("🧾 검증 / 대조 🧾")

        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        if st.button("세금계산서 대조"):
            go("tax")

        st.button("사학진흥재단 차입금 정리", disabled=True)
        st.button("선급법인세 취합", disabled=True)
        st.markdown("</div>", unsafe_allow_html=True)


def main():
    st.set_page_config(layout="wide", page_title="재무·세무 자동화 메인")

    if "page" not in st.session_state:
        st.session_state["page"] = "main"

    if st.session_state["page"] == "main":
        render_main_menu()

    elif st.session_state["page"] == "tax":
        # 상단에 뒤로가기 버튼 하나 붙이기
        back_col, title_col = st.columns([1, 5])
        with back_col:
            if st.button("⬅ 메인으로"):
                go("main")
        with title_col:
            st.empty()  # run_tax 안에서 제목을 찍을 거라면 비워둬도 됨

        # 분리해 둔 세금계산서 기능 실행
        run_tax()

    elif st.session_state["page"] == "misc":
        # 기타 기능 페이지 (misc_app.run)
        run_misc()


if __name__ == "__main__":
    main()
