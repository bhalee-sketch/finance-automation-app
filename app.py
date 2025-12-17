# app.py
# -*- coding: utf-8 -*-
import streamlit as st
from tax_invoice_app import run as run_tax  # ← 분리한 파일에서 run() 가져오기
from misc_app import run as run_misc        # ← 기타기능 모듈
from loan_app import run as run_loan
from ledger_app import run as run_ledger
from xls_convert_app import run as run_xls_convert
from fundcheck_app import run as run_fund_check
from donation_main_app import run as run_donation_main
from expense_account_check_app import run as run_expense_account_check
from prepaid_cit_app import run as run_prepaid_cit

def go(page: str):
    """페이지 상태 변경 + 즉시 리렌더링"""
    st.session_state["page"] = page
    st.rerun()

def render_main_menu():

    # 🔐 숨김 설명서 (서버 관리 / 이용법)
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

    col1, col2, col3 = st.columns(3)

    # ---------------------- 결산 작업 + 기타기능 ----------------------
    with col1:
        st.subheader("📘 결산 작업 📘")

        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        st.button("재무제표 생성", disabled=True)
        if st.button("회계단위별 원장파일 통합"):
            go("ledger")
        st.button("재무제표 vs 부속명세서 검증", disabled=True)

        st.markdown("---")   # 구분선

        st.subheader("🛠️ 기타기능 🛠️")
        if st.button("자금이체 적요 자동생성"):
            go("misc")
        if st.button("XLS → XLSX 변환"):
            go("xls_convert")    
        st.markdown("</div>", unsafe_allow_html=True)

    # ---------------------- 검증 / 대조 ----------------------
    with col2:
        st.subheader("🧾 검증 / 대조 🧾")

        st.markdown('<div class="small-button">', unsafe_allow_html=True)
        if st.button("세금계산서 대조"):
            go("tax")
        if st.button("사학진흥재단 차입금 정리"):
            go("loan")     # loan_app.py를 연결할 key
        if st.button("선급법인세 취합"):
            go("prepaid_cit")
        if st.button("지출계좌 재원 검증"):
            go("expense_account_check")
        if st.button("임의기금 지출계좌 검증"):
            go("fund_check")
        if st.button("출연받은재산 정리"):
            go("donation_main")

    # ---------------------- 출연받은 재산 작업 ----------------------
    with col3:
        st.subheader("🎁출연받은재산 보고를 위한 작업🎁")
        st.write("아래의 기능들을 순서대로 작업하는 것을 추천")    
        st.button("1) 당해 기부금 내역 정리", disabled=True)
        st.button("2) 출연받은재산보고 정리", disabled=True) 
        st.button("3) 기부금지출명세서 정리", disabled=True)
        st.button("4) 기부금지출명세서 검증", disabled=True)    
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("---")   # 구분선

        st.subheader("산단 준비중")
        
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

    elif st.session_state["page"] == "loan":
        run_loan()

    elif st.session_state["page"] == "ledger":
        run_ledger()

    elif st.session_state["page"] == "xls_convert":
        run_xls_convert()
    
    elif st.session_state["page"] == "fund_check":
        run_fund_check()

    elif st.session_state["page"] == "donation_main":
        run_donation_main()

    elif st.session_state["page"] == "expense_account_check":
        run_expense_account_check()
        
    elif st.session_state["page"] == "prepaid_cit":
        run_prepaid_cit()

if __name__ == "__main__":
    main()
