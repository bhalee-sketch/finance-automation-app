# pages/utils.py
# -*- coding: utf-8 -*-

from __future__ import annotations

from pathlib import Path
from io import BytesIO
import re

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parents[1]   # 전자자료/
DATA_DIR = BASE_DIR / "data"
PLACEHOLDER_MENUS = ["연도별 증감현황", "연도별 구성현황", "분석(미정) 3", "분석(미정) 4"]

def list_data_files():
    if not DATA_DIR.exists():
        return []
    files = sorted(DATA_DIR.glob("*.xlsx"))
    files += sorted(DATA_DIR.glob("*.xlsm"))
    return files

def year_from_filename(stem: str) -> str:
    """
    파일명(확장자 제외)에서 연도(20XX) 추출.
    예: '2024회계연도' -> '2024'
    없으면 stem 그대로 반환
    """
    m = re.search(r"(20\d{2})", str(stem))
    return m.group(1) if m else str(stem)


def safe_numeric(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace({"nan": None, "None": None, "": None})
        .pipe(pd.to_numeric, errors="coerce")
    )


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "data") -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def top_menu():
    if "page" not in st.session_state:
        st.session_state["page"] = "재무제표"

    menus = ["재무제표"] + PLACEHOLDER_MENUS

    st.markdown(
        """
        <style>
        .topmenu-wrap {
            position: sticky;
            top: 0;
            z-index: 999;
            background: white;
            padding: 0.6rem 0 0.2rem 0;
            border-bottom: 1px solid rgba(49, 51, 63, 0.1);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="topmenu-wrap">', unsafe_allow_html=True)
    cols = st.columns(len(menus))
    for i, m in enumerate(menus):
        with cols[i]:
            is_active = (st.session_state["page"] == m)
            label = f"✅ {m}" if is_active else m
            if st.button(label, use_container_width=True):
                st.session_state["page"] = m
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
