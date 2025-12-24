# pages/raw.py
# -*- coding: utf-8 -*-

from __future__ import annotations

import re
import pandas as pd
import streamlit as st
import plotly.express as px

from statement.pages.utils import list_data_files, year_from_filename, to_excel_bytes, safe_numeric

# =========================
# 시트명 파싱: "자금계산서(전체)" 같은 규칙
# =========================
SHEET_PATTERN = re.compile(
    r"^\s*(자금계산서|재무상태표|운영계산서)\s*\(\s*(전체|등록금|비등록금)\s*\)\s*$"
)


def parse_statement_sheets(sheet_names: list[str]) -> dict[tuple[str, str], str]:
    """
    반환: {(제표, 구분): 실제 시트명}
    예: {("자금계산서","전체"): "자금계산서(전체)", ...}
    """
    mapping: dict[tuple[str, str], str] = {}
    for name in sheet_names:
        m = SHEET_PATTERN.match(str(name))
        if m:
            stmt, unit = m.group(1), m.group(2)
            mapping[(stmt, unit)] = name
    return mapping


# =========================
# 자금계산서: 블록 기준 수입/지출 분리
# =========================
def classify_cashflow_by_blocks(subjects: pd.Series) -> pd.Series:
    START_IN = "등록금및수강료수입"
    END_IN = "자금수입총계"
    START_OUT = "보수"
    END_OUT = "자금지출총계"

    state = "기타"
    out = []

    for v in subjects.astype(str).fillna(""):
        s = v.strip().replace("\u00a0", " ")

        if START_IN in s:
            state = "수입"
        if START_OUT in s:
            state = "지출"

        out.append(state)

        # 종료행은 포함하고, 다음 행부터 기타로
        if END_IN in s and state == "수입":
            state = "기타"
        if END_OUT in s and state == "지출":
            state = "기타"

    return pd.Series(out, index=subjects.index)

# =========================
# 재무상태서: 블록 기준 수입/지출 분리
# =========================
def classify_bs_assets_liab_equity(subjects: pd.Series) -> pd.Series:
    START_ASSET = "유동자산"
    END_ASSET = "자산총계"
    START_LIAB = "유동부채"
    # 학교 결산서 표현이 조금씩 달라서 후보를 넉넉히
    END_LIAB_CANDIDATES = ["부채와기본금총계", "부채및기본금총계", "기본금총계"]

    state = "기타"
    out = []

    for v in subjects.astype(str).fillna(""):
        s = v.strip().replace("\u00a0", " ")

        if START_ASSET in s:
            state = "자산"
        if START_LIAB in s:
            state = "부채/기본금"

        out.append(state)

        # 종료행은 포함하고 다음부터 기타
        if END_ASSET in s and state == "자산":
            state = "기타"
        if state == "부채/기본금" and any(x in s for x in END_LIAB_CANDIDATES):
            state = "기타"

    return pd.Series(out, index=subjects.index)
# 일반 시트용(임시) 키워드 분류
def _classify_income_expense(subject: str) -> str:
    s = str(subject)
    income_kw = ["수입", "수익", "등록금", "기부금", "전입금", "보조금", "수강료", "이자수입", "잡수입"]
    expense_kw = ["지출", "비용", "경비", "급여", "수당", "장학", "연구", "시설", "공사", "감가상각", "상각", "이자비용", "잡손실"]
    if any(k in s for k in income_kw):
        return "수입"
    if any(k in s for k in expense_kw):
        return "지출"
    return "기타"


# =========================
# 표 가독성 개선: 들여쓰기 레벨/콤마
# =========================
def _indent_level(text: str) -> int:
    """앞 공백(스페이스) 5칸 = 1레벨"""
    if text is None:
        return 0
    s = str(text)

    # NBSP( )도 스페이스로 통일
    s = s.replace("\u00a0", " ")

    # 앞쪽 공백 개수만 세기 (탭은 4칸으로 간주)
    leading_spaces = 0
    for ch in s:
        if ch == " ":
            leading_spaces += 1
        elif ch == "\t":
            leading_spaces += 4
        else:
            break

    return leading_spaces // 5

def _is_expense_separator(text: str) -> bool:
    """[지출]----- 같은 구분행 감지"""
    if text is None:
        return False
    s = str(text).replace("\u00a0", " ")
    return bool(re.search(r"[\[\［【]\s*지출\s*[\]\］】]\s*[-=—–]{3,}", s))

def calc_df_height(n_rows: int, row_h: int = 34, header_h: int = 38, padding: int = 16) -> int:
    """
    dataframe 내부 스크롤 제거용 높이 계산
    """
    return header_h + n_rows * row_h + padding

def prettify_raw_table(raw: pd.DataFrame):
    df = raw.copy()

    if "과목" not in df.columns:
        raise ValueError("현재 시트에 '과목' 컬럼이 없습니다.")

    # index 제거
    df = df.reset_index(drop=True)

    # 과목 문자열화 + 들여쓰기 레벨
    df["과목"] = df["과목"].astype(str)
    level_series = df["과목"].map(_indent_level)

    # ✅ [지출]----- 행을 '원본 과목' 기준으로 먼저 잡아두기 (이게 핵심)
    sep_rows = df["과목"].apply(_is_expense_separator)

    # money 컬럼 숫자 변환
    money_cols = [c for c in df.columns if c not in ["과목", "Rate"]]
    for c in money_cols:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False),
            errors="coerce",
        )
    # ✅ 구분행 마스크 (원본 과목 기준)
    sep_rows = df["과목"].apply(_is_expense_separator)

    # money_cols / Rate 숫자 변환 끝난 뒤에 "구분행만" 비우기 (dtype 유지)
    if sep_rows.any():
        df.loc[sep_rows, money_cols] = pd.NA
        if "Rate" in df.columns:
            df.loc[sep_rows, "Rate"] = pd.NA
        df.loc[sep_rows, "과목"] = " "   # 행 높이 유지용

    if "Rate" in df.columns:
        df["Rate"] = (
            df["Rate"].astype(str).str.replace("%", "", regex=False).pipe(safe_numeric)
        )

    # ✅ 구분행은 화면에서 값이 안 보이게 만들기(표 안 “띠”)
    if sep_rows.any():
        df.loc[sep_rows, "과목"] = " "    # 과목은 공백 1칸(행 높이 유지)

    subj_idx = list(df.columns).index("과목")

    def _row_css_by_level(idx: int):
        lvl = int(level_series.iloc[idx])

        # ✅ 지출 구분 행: 배경 #F2F2F2, 글자색도 #F2F2F2(완전 숨김)
        if bool(sep_rows.iloc[idx]):
            return [
                "background-color:#F2F2F2 !important; color:#F2F2F2 !important; font-weight:900;"
            ] * len(df.columns)

        # 관(레벨0)
        if lvl == 0:
            css = ["background-color:#2b1d1d; color:#f1f3f5; font-weight:700;"] * len(df.columns)
            css[subj_idx] = "background-color:#2b1d1d; color:#ffffff; font-weight:900;"
            return css

        # 항(레벨1)
        if lvl == 1:
            css = ["background-color:#24282e; color:#f1f3f5;"] * len(df.columns)
            css[subj_idx] = "background-color:#24282e; color:#ffffff; font-weight:800;"
            return css

        # 목(레벨2+)
        css = [""] * len(df.columns)
        css[subj_idx] = "font-weight:600; opacity:0.85;"
        return css

    styler = df.style.apply(lambda row: _row_css_by_level(row.name), axis=1)

    fmt = {c: "{:,.0f}" for c in money_cols}
    if "Rate" in df.columns:
        fmt["Rate"] = "{:,.1f}"
    styler = styler.format(fmt, na_rep="")

    return styler


# =========================
# (옵션) 롱포맷 미리보기용
# =========================
def find_subject_col(df: pd.DataFrame) -> str:
    candidates = ["과목", "계정", "항목", "과목명", "계정과목", "계정명"]
    for c in df.columns:
        if str(c).strip() in candidates:
            return c
    for c in df.columns:
        txt = str(c)
        if any(k in txt for k in candidates):
            return c
    raise ValueError("과목(계정/항목) 컬럼을 찾지 못했습니다.")

def tidy_from_sheet(df: pd.DataFrame, year: int) -> pd.DataFrame:
    subject_col = find_subject_col(df)

    fund_cols = [c for c in ["등록금", "비등록금", "내부", "확정", "전용", "예비비"] if c in df.columns]
    result_cols = [c for c in ["예산", "결산", "증감"] if c in df.columns]
    rate_col = "Rate" if "Rate" in df.columns else None

    base = df.copy()
    base[subject_col] = base[subject_col].astype(str).str.strip()
    base = base[base[subject_col].notna() & (base[subject_col] != "")]

    parts = []

    if fund_cols:
        melted_fund = base.melt(
            id_vars=[subject_col],
            value_vars=fund_cols,
            var_name="재원구분",
            value_name="금액",
        )
        melted_fund["연도"] = year
        melted_fund["금액유형"] = "결산"
        melted_fund["금액"] = safe_numeric(melted_fund["금액"]).fillna(0)
        melted_fund.rename(columns={subject_col: "과목"}, inplace=True)
        parts.append(melted_fund)

    if result_cols:
        melted_result = base.melt(
            id_vars=[subject_col],
            value_vars=result_cols,
            var_name="금액유형",
            value_name="금액",
        )
        melted_result["연도"] = year
        melted_result["재원구분"] = "전체"
        melted_result["금액"] = safe_numeric(melted_result["금액"]).fillna(0)
        melted_result.rename(columns={subject_col: "과목"}, inplace=True)
        parts.append(melted_result)

    if rate_col:
        rate_df = base[[subject_col, rate_col]].copy()
        rate_df.rename(columns={subject_col: "과목", rate_col: "금액"}, inplace=True)
        rate_df["연도"] = year
        rate_df["재원구분"] = "전체"
        rate_df["금액유형"] = "Rate"
        rate_df["금액"] = (
            rate_df["금액"].astype(str).str.replace("%", "", regex=False).pipe(safe_numeric)
        )
        parts.append(rate_df)

    if not parts:
        raise ValueError("변환할 수 있는 컬럼을 찾지 못했습니다.")

    long_df = pd.concat(parts, ignore_index=True)
    long_df["절대값"] = long_df["금액"].abs()
    return long_df

# =========================
# 페이지 렌더
# =========================
def render():
    st.title("📄 재무제표 📄")

    # ✅ 여기다가 넣으세요 (CSS는 한 번만)
    st.markdown(
        """
        <style>
        /* selectbox 전체 클릭 영역 */
        div[data-baseweb="select"] { cursor: pointer !important; }
        div[data-baseweb="select"] * { cursor: pointer !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    files = list_data_files()
    if not files:
        st.error("data/ 폴더에 엑셀 파일이 없습니다. 예: data/2024회계연도.xlsx")
        return

    # ✅ 회계연도
    file_options = []
    for p in files:
        yr = year_from_filename(p.stem)
        file_options.append((yr, p)) 
    # 연도 목록 (문자 → 정렬용)
    years = [int(x[0]) for x in file_options]
    latest_year = max(years)
    c_year, _ = st.columns([1, 6])  # 왼쪽만 좁게
    with c_year:
        sel_label = st.selectbox(
            "회계연도",
            [x[0] for x in file_options],
            index=years.index(latest_year),  # ✅ 핵심
            key="year"
        )

    sel_path = dict(file_options)[sel_label]

    # 시트 파싱
    xls = pd.ExcelFile(sel_path)
    sheet_map = parse_statement_sheets(xls.sheet_names)

    if not sheet_map:
        st.error(
            "시트명 규칙을 찾지 못했습니다.\n\n"
            "예: 자금계산서(전체), 자금계산서(등록금), 자금계산서(비등록금)\n"
            "또는 재무상태표(전체) 형태로 시트명을 맞춰주세요."
        )
        return

    # ✅ 제표 / 구분 선택 (라디오 버튼 UI)
    col1, col2 = st.columns(2)

    with col1:
        statement_type = st.radio(
            "제표 선택",
            ["자금계산서", "재무상태표", "운영계산서"],
            index=0,
            label_visibility="collapsed",
            key="statement_type",
        )
    with col2:
        UNIT_LABELS = {
            "교비전체": "전체",
            "교비 - 등록금": "등록금",
            "교비 - 비등록금": "비등록금",
        }
        unit_label = st.radio(
            "구분 선택",
            list(UNIT_LABELS.keys()),
            index=0,
            label_visibility="collapsed",
            key="unit_label",
        )

    unit_type = UNIT_LABELS[unit_label]

    sheet = sheet_map[(statement_type, unit_type)]
    
    # 원본 읽기
    raw = pd.read_excel(sel_path, sheet_name=sheet)
    st.caption(f"파일: {sel_path.name} / 시트: {sheet} / 행 {len(raw):,} / 열 {raw.shape[1]:,}")

    # 분류(자금계산서는 블록 규칙)
    df_base = raw.copy()
    if "과목" not in df_base.columns:
        st.error("현재 시트에 '과목' 컬럼이 없습니다. (헤더명을 확인해주세요)")
        return

    # 분류
    if statement_type == "자금계산서":
        df_base["_구분"] = classify_cashflow_by_blocks(df_base["과목"])
        tabs = ["전체", "수입", "지출"]
    elif statement_type == "재무상태표":
        df_base["_구분"] = classify_bs_assets_liab_equity(df_base["과목"])
        tabs = ["전체", "자산", "부채/기본금"]
    elif statement_type == "운영계산서":
        df_base["_구분"] = classify_cashflow_by_blocks(df_base["과목"])
        tabs = ["전체", "수입", "지출"]
    
    tab_all, tab_1, tab_2 = st.tabs(tabs)

    def calc_df_height(n_rows: int, row_h: int = 35, header_h: int = 38, padding: int = 16) -> int:
        # Streamlit dataframe 행 높이가 대략 35px 전후라서 이 정도로 맞추면 스크롤이 거의 사라짐
        return header_h + n_rows * row_h + padding

    with tab_all:
        show = df_base.drop(columns=["_구분"], errors="ignore")
        st.dataframe(
            prettify_raw_table(show),
            use_container_width=True,
            height=calc_df_height(len(show))
        )

    with tab_1:
        key = "자산" if statement_type == "재무상태표" else "수입"
        d1 = df_base[
            (df_base["_구분"] == key)
            & (~df_base["과목"].apply(_is_expense_separator))
        ].drop(columns=["_구분"], errors="ignore")

        st.dataframe(
            prettify_raw_table(d1),
            use_container_width=True,
            height=calc_df_height(len(d1)),
        )

    with tab_2:
        key = "부채/기본금" if statement_type == "재무상태표" else "지출"
        d2 = df_base[
            df_base["_구분"] == key
        ].drop(columns=["_구분"], errors="ignore")

        st.dataframe(
            prettify_raw_table(d2),
            use_container_width=True,
            height=calc_df_height(len(d2)),
        )

    st.download_button(
        "⬇️ 현재 시트를 그대로 엑셀로 다운로드",
        data=to_excel_bytes(raw, sheet_name="raw"),
        file_name=f"원본_{sel_path.stem}_{sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()