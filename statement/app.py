# statement/app.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st

from statement.pages.raw import render as render_raw
from statement.pages.analysis1 import render as render_analysis1
from statement.pages.analysis2 import render as render_analysis2
from statement.pages.placeholder import render as render_placeholder

FS_MENUS = ["재무제표", "연도별 증감현황", "연도별 구성현황"]

def render(go):
    left, mid, right = st.columns([1, 3, 2])

    with left:
        if st.button("⬅ 홈", key="fs_back_home"):
            go("home")

    with right:
        if "fs_page" not in st.session_state:
            st.session_state["fs_page"] = "재무제표"
        st.session_state["fs_page"] = st.selectbox(
            "메뉴",
            FS_MENUS,
            index=FS_MENUS.index(st.session_state["fs_page"])
            if st.session_state["fs_page"] in FS_MENUS else 0,
        )

    page = st.session_state["fs_page"]

    if page == "재무제표":
        render_raw()
    elif page == "연도별 증감현황":
        render_analysis1()
    elif page == "연도별 구성현황":
        render_analysis2()
    else:
        render_placeholder(page)
