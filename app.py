# app.py (프로젝트 루트)
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st

from statement.app import render as render_statement
from excel.app import render as render_excel

def go(page: str):
    st.session_state["page"] = page
    st.rerun()

def render_home():
    st.title("🏠 재무회계팀 통합 시스템")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📈 재무제표 현황")
        if st.button("재무제표 현황", use_container_width=True):
            go("FS:")

    with c2:
        st.subheader("🧰 엑셀 정리 작업")
        if st.button("엑셀 정리 작업", use_container_width=True):
            go("EXCEL:")

def main():
    st.set_page_config(layout="wide", page_title="재무회계팀 통합 시스템")

    if "page" not in st.session_state:
        st.session_state["page"] = "home"

    page = st.session_state["page"]

    if page == "home":
        render_home()
    elif page.startswith("FS:"):
        render_statement(go=go)     # ✅ statement 쪽으로 진입
    elif page.startswith("EXCEL:"):
        render_excel(go=go)         # ✅ app(엑셀) 쪽으로 진입
    else:
        go("home")

if __name__ == "__main__":
    main()
