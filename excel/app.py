# app/app.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st
from excel.tax_invoice_app import run as run_tax
from excel.misc_app import run as run_misc
from excel.loan_app import run as run_loan
from excel.ledger_app import run as run_ledger
from excel.xls_convert_app import run as run_xls_convert
from excel.fundcheck_app import run as run_fund_check
from excel.donation_main_app import run as run_donation_main
from excel.expense_account_check_app import run as run_expense_account_check
from excel.prepaid_cit_app import run as run_prepaid_cit


def render_main_menu(go):
    # 🔐 숨김 설명서
    with st.expander("🛠 서버 관리 · 이용 방법 (클릭해서 열기)", expanded=False):
        st.markdown(
            """
            ### 📌 기본 안내
            - 본 시스템은 **재무회계팀 내부 전용** 자동화 도구입니다.
            - 크롬(Chrome) 브라우저 사용을 권장합니다.
            - 엑셀 파일 업로드 시 **파일명·시트 구조 변경 금지**.

            ### 🖥 버전 업데이트 방법
            - https://github.com 접속
            - Add file → Upload files 클릭 후 코딩한 app.py파일 업로드
            - 맨 아래 Commit changes 클릭

            ### ⚠ 주의사항
            - 업로드한 파일은 **서버에 저장되지 않습니다**.
            - 개인정보 포함 파일은 작업 후 즉시 삭제 권장.
            - 동시에 여러 기능을 새 탭에서 실행하지 마세요.
            """
        )

    st.title("📊 재무회계팀 자동화 작업 메뉴")
    st.write("원하는 작업을 선택하세요.")

    st.markdown(
        """
        <style>
            .small-button button { width: 150px !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("📘 결산 작업 📘")
        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        st.button("재무제표 생성", disabled=True)
        if st.button("회계단위별 원장파일 통합"):
            go("EXCEL:ledger")
        st.button("재무제표 vs 부속명세서 검증", disabled=True)

        st.markdown("---")
        st.subheader("🛠️ 기타기능 🛠️")
        if st.button("자금이체 적요 자동조성"):
            go("EXCEL:misc")
        if st.button("XLS → XLSX 변환"):
            go("EXCEL:xls_convert")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.subheader("🧾 검증 / 대조 🧾")
        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        if st.button("세금계산서 대조"):
            go("EXCEL:tax")
        if st.button("사학진흥재단 차입금 정리"):
            go("EXCEL:loan")
        if st.button("선급법인세 취합"):
            go("EXCEL:prepaid_cit")
        if st.button("지출계좌 재원 검증"):
            go("EXCEL:expense_account_check")
        if st.button("임의기금 지출계좌 검증"):
            go("EXCEL:fund_check")
        if st.button("출연받은재산 정리"):
            go("EXCEL:donation_main")

    with col3:
        st.subheader("🎁출연받은재산 보고를 위한 작업🎁")
        st.write("아래의 기능들을 순서대로 작업하는 것을 추천")
        st.button("1) 당해 기부금 내역 정리", disabled=True)
        st.button("2) 출연받은재산보고 정리", disabled=True)
        st.button("3) 기부금지출명세서 정리", disabled=True)
        st.button("4) 기부금지출명세서 검증", disabled=True)
        st.markdown("---")
        st.subheader("산단 준비중")


def render(go):
    # ✅ 홈 버튼(통합 메인으로)
    if st.button("⬅ 홈", key="excel_back_home"):
        go("home")

    # ✅ 엑셀 내부 페이지 키는 따로 (메인 page와 충돌 방지)
    if "excel_page" not in st.session_state:
        st.session_state["excel_page"] = "EXCEL:main"

    page = st.session_state.get("page", "EXCEL:main")  # 통합 메인이 내려준 값 사용

    if page == "EXCEL:main":
        render_main_menu(go)

    elif page == "EXCEL:tax":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_tax"):
            go("EXCEL:main")
        run_tax()

    elif page == "EXCEL:misc":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_misc"):
            go("EXCEL:main")
        run_misc()

    elif page == "EXCEL:loan":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_loan"):
            go("EXCEL:main")
        run_loan()

    elif page == "EXCEL:ledger":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_ledger"):
            go("EXCEL:main")
        run_ledger()

    elif page == "EXCEL:xls_convert":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_xls"):
            go("EXCEL:main")
        run_xls_convert()

    elif page == "EXCEL:fund_check":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_fund"):
            go("EXCEL:main")
        run_fund_check()

    elif page == "EXCEL:donation_main":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_donation"):
            go("EXCEL:main")
        run_donation_main()

    elif page == "EXCEL:expense_account_check":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_expense"):
            go("EXCEL:main")
        run_expense_account_check()

    elif page == "EXCEL:prepaid_cit":
        if st.button("⬅ 엑셀메뉴", key="back_excel_menu_prepaid"):
            go("EXCEL:main")
        run_prepaid_cit()

    else:
        go("EXCEL:main")
