# pages/analysis1.py
# -*- coding: utf-8 -*-

from __future__ import annotations

import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from statement.pages.analysis1_config import IO_GUAN_GROUP, BS_GUAN_GROUP
from statement.pages.utils import list_data_files, year_from_filename, safe_numeric


# ✅ 커스텀 옵션 완전 비활성화
CUSTOM_OPTIONS = {}
def series_cashsheet_row_total(unit_type: str, value_col: str) -> pd.DataFrame:
    """
    자금계산서에서 '자 금 지 출 총 계' 행을 찾아 value_col(결산)을 연도별로 가져옴
    """
    files = list_data_files()
    rows = []

    for p in files:
        year_txt = year_from_filename(p.stem)
        try:
            year = int(year_txt)
        except Exception:
            continue

        try:
            path_str = str(p)
            sheet_map = _cached_sheet_map(path_str)
            sheet = sheet_map.get(("자금계산서", unit_type))
            if not sheet:
                continue

            df = _cached_read_sheet(path_str, sheet)
            subj = find_subject_col(df)

            if value_col not in df.columns:
                continue

            # 과목명 정규화
            subj_norm = (
                df[subj].astype(str)
                .str.replace("\u00a0", " ", regex=False)
                .map(_norm)
            )

            # ✅ 타겟 행 찾기
            hit_idx = subj_norm[subj_norm == TARGET_TOTAL_LABEL_NORM].index
            if len(hit_idx) == 0:
                continue

            idx = int(hit_idx[-1])  # 혹시 여러 개면 마지막
            val = safe_numeric(df.loc[[idx], value_col]).iloc[0]
            if pd.isna(val):
                continue

            rows.append({"연도": year, "금액": float(val)})

        except Exception:
            continue

    if not rows:
        return pd.DataFrame(columns=["연도", "금액"])

    return pd.DataFrame(rows).sort_values("연도")

# ======================================================
# 캐시: Excel 반복 읽기 방지
# ======================================================
@st.cache_data(show_spinner=False)
def _cached_sheet_map(path_str: str) -> dict[tuple[str, str], str]:
    xls = pd.ExcelFile(path_str)
    return parse_statement_sheets(xls.sheet_names)


@st.cache_data(show_spinner=False)
def _cached_read_sheet(path_str: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(path_str, sheet_name=sheet_name)

def series_cashsheet_last_row_total(unit_type: str, value_col: str) -> pd.DataFrame:
    """
    자금계산서(단위: 전체/등록금/비등록금)에서
    '맨 아래(마지막 행)'의 value_col(결산)을 연도별로 가져와 총계로 사용
    """
    files = list_data_files()
    rows = []

    for p in files:
        year_txt = year_from_filename(p.stem)
        try:
            year = int(year_txt)
        except Exception:
            continue

        try:
            path_str = str(p)
            sheet_map = _cached_sheet_map(path_str)
            sheet = sheet_map.get(("자금계산서", unit_type))
            if not sheet:
                continue

            df = _cached_read_sheet(path_str, sheet)
            if value_col not in df.columns:
                continue

            vals = safe_numeric(df[value_col])

            # ✅ 마지막 유효 값(빈칸/NaN 제외)
            last = vals.dropna()
            if last.empty:
                continue

            last_val = float(last.iloc[-1])
            rows.append({"연도": year, "금액": last_val})

        except Exception:
            continue

    if not rows:
        return pd.DataFrame(columns=["연도", "금액"])

    out = pd.DataFrame(rows).sort_values("연도")
    return out

# ======================================================
# 들여쓰기(스페이스) 기반 관/항/목
# - 관: 0
# - 항: 5
# - 목: 10 (이상)
# ======================================================
def _leading_spaces(text: str) -> int:
    if text is None:
        return 0
    s = str(text).replace("\u00a0", " ")
    n = 0
    for ch in s:
        if ch == " ":
            n += 1
        elif ch == "\t":
            n += 4
        else:
            break
    return n

def depth_rules(statement_type: str) -> Tuple[int, int, int]:
    return (0, 5, 10)

def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s).replace("\u00a0", " ")).strip()

TARGET_TOTAL_LABEL_NORM = _norm("자 금 지 출 총 계")  # => "자금지출총계"

# ======================================================
# 시트명 파서
# ======================================================
SHEET_PATTERN = re.compile(r"^\s*(자금계산서|재무상태표|운영계산서)\s*\(\s*(전체|등록금|비등록금)\s*\)\s*$")

def parse_statement_sheets(sheet_names: List[str]) -> Dict[Tuple[str, str], str]:
    mapping: Dict[Tuple[str, str], str] = {}
    for name in sheet_names:
        m = SHEET_PATTERN.match(str(name))
        if m:
            stmt, unit = m.group(1), m.group(2)
            mapping[(stmt, unit)] = name
    return mapping

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

# ======================================================
# 수입/지출(또는 자산/부채/기본금) 분류
# ======================================================
@st.cache_data(show_spinner=False)
def _classify_io(statement_type: str, guan: str) -> str:
    g = (guan or "").strip()

    if statement_type in ("자금계산서", "운영계산서"):
        if g in IO_GUAN_GROUP:
            return IO_GUAN_GROUP[g]
    elif statement_type == "재무상태표":
        if g in BS_GUAN_GROUP:
            return BS_GUAN_GROUP[g]

    gn = g.replace(" ", "")
    income_kw = ["수입", "수익", "전입금", "기부금", "보조금", "등록금", "교육부대수익", "운영수익"]
    expense_kw = ["지출", "비용", "운영비용", "관리운영비", "교육비", "연구비", "장학금", "일반관리비"]

    if any(k.replace(" ", "") in gn for k in income_kw):
        return "수입"
    if any(k.replace(" ", "") in gn for k in expense_kw):
        return "지출"
    return "기타"
# 재무상태표 합성 관 정의
ASSET_TOTAL_GUANS = {
    _norm("유동자산"),
    _norm("투자와기타자산"),
    _norm("고정자산"),
}

LIABILITY_TOTAL_GUANS = {
    _norm("유동부채"),
    _norm("고정부채"),
}
# ======================================================
# 최신 파일(최신연도)에서 관/목 순서 추출 → 드롭다운 순서 안정화
# ======================================================
def _latest_file_path():
    files = list_data_files()
    pairs = []
    for p in files:
        ytxt = year_from_filename(p.stem)
        try:
            y = int(ytxt)
        except Exception:
            continue
        pairs.append((y, p))
    if not pairs:
        return None
    pairs.sort(key=lambda x: x[0], reverse=True)
    return str(pairs[0][1])

def get_guan_order_from_files(statement_type: str, unit_type: str) -> List[str]:
    guan_d, _, _ = depth_rules(statement_type)
    path_str = _latest_file_path()
    if not path_str:
        return []

    try:
        sheet_map = _cached_sheet_map(path_str)
        sheet = sheet_map.get((statement_type, unit_type))
        if not sheet:
            return []

        df = _cached_read_sheet(path_str, sheet)
        subj = find_subject_col(df)

        subjects_raw = (
            df[subj].astype(str)
            .str.replace("\u00a0", " ", regex=False)
            .str.rstrip()
        )

        seen = set()
        out: List[str] = []
        for raw in subjects_raw.tolist():
            if str(raw).strip() == "":
                continue
            if _leading_spaces(raw) == guan_d:
                name = str(raw).strip()
                key = _norm(name)
                if key and key not in seen:
                    seen.add(key)
                    out.append(name)
        return out
    except Exception:
        return []

def get_hang_order_from_files(statement_type: str, unit_type: str) -> List[str]:
    _, hang_d, mok_d = depth_rules(statement_type)
    path_str = _latest_file_path()
    if not path_str:
        return []

    try:
        sheet_map = _cached_sheet_map(path_str)
        sheet = sheet_map.get((statement_type, unit_type))
        if not sheet:
            return []

        df = _cached_read_sheet(path_str, sheet)
        subj = find_subject_col(df)

        subjects_raw = (
            df[subj].astype(str)
            .str.replace("\u00a0", " ", regex=False)
            .str.rstrip()
        )

        seen = set()
        out: List[str] = []
        for raw in subjects_raw.tolist():
            if str(raw).strip() == "":
                continue

            depth = _leading_spaces(raw)
            name = str(raw).strip()

            # ✅ 항(= 정확히 hang depth)만
            if depth == hang_d:
                key = _norm(name)
                if key and key not in seen:
                    seen.add(key)
                    out.append(name)

            # (옵션) 목으로 내려가면 항 수집 로직에 영향 없지만,
            #       굳이 끊고 싶으면 아래처럼 써도 됩니다.
            # if depth >= mok_d: 
            #     continue

        return out
    except Exception:
        return []

def get_mok_order_from_files(statement_type: str, unit_type: str) -> List[str]:
    _, _, mok_d = depth_rules(statement_type)
    path_str = _latest_file_path()
    if not path_str:
        return []

    try:
        sheet_map = _cached_sheet_map(path_str)
        sheet = sheet_map.get((statement_type, unit_type))
        if not sheet:
            return []

        df = _cached_read_sheet(path_str, sheet)
        subj = find_subject_col(df)

        subjects_raw = (
            df[subj].astype(str)
            .str.replace("\u00a0", " ", regex=False)
            .str.rstrip()
        )

        seen = set()
        out: List[str] = []
        for raw in subjects_raw.tolist():
            if str(raw).strip() == "":
                continue
            if _leading_spaces(raw) >= mok_d:
                n = _norm(str(raw).strip())
                if n and n not in seen:
                    seen.add(n)
                    out.append(n)  # norm 저장
        return out
    except Exception:
        return []

# ======================================================
# 시계열 구축
# - 기본: 목 라인만 집계
# - 예외(요청사항): "미사용전기이월자금", "미사용차기이월자금"은 "관 헤더행 값" 그대로
#   => 관 라인에서 mok를 관명으로 채워서 살아남게 처리
# ======================================================
SPECIAL_GUAN_DIRECT = {"미사용전기이월자금", "미사용차기이월자금"}

def build_timeseries(statement_type: str, unit_type: str, value_col: str) -> pd.DataFrame:
    guan_d, hang_d, mok_d = depth_rules(statement_type)
    files = list_data_files()
    rows = []

    for p in files:
        year_txt = year_from_filename(p.stem)
        try:
            year = int(year_txt)
        except Exception:
            continue

        try:
            path_str = str(p)
            sheet_map = _cached_sheet_map(path_str)
            sheet = sheet_map.get((statement_type, unit_type))
            if not sheet:
                continue

            df = _cached_read_sheet(path_str, sheet)
            subj = find_subject_col(df)

            if value_col not in df.columns:
                continue

            vals = safe_numeric(df[value_col]).fillna(0)

            subjects_raw = (
                df[subj].astype(str)
                .str.replace("\u00a0", " ", regex=False)
                .str.rstrip()
            )

            tmp = pd.DataFrame({"연도": year, "과목_raw": subjects_raw, "금액": vals})
            tmp = tmp[tmp["과목_raw"].notna() & (tmp["과목_raw"] != "")].copy()

            guan = ""
            hang = ""
            mok = ""
            guan_list, hang_list, mok_list = [], [], []

            for s in tmp["과목_raw"].tolist():
                depth = _leading_spaces(s)
                name = str(s).strip()

                if depth == guan_d:
                    guan = name
                    hang = ""
                    mok = ""

                    # ✅ 특수 관은 관 헤더행 자체를 데이터로 취급
                    if name in SPECIAL_GUAN_DIRECT:
                        mok = name

                elif depth == hang_d:
                    hang = name
                    mok = ""

                elif depth >= mok_d:
                    mok = name

                guan_list.append(guan)
                hang_list.append(hang)
                mok_list.append(mok)

            tmp["관"] = guan_list
            tmp["항"] = hang_list
            tmp["목"] = mok_list
            tmp["구분"] = tmp["관"].map(lambda x: _classify_io(statement_type, x))

            # ✅ 목이 빈 행 제거 (특수 관 헤더행은 목이 채워져서 살아남음)
            tmp = tmp[tmp["목"].astype(str).str.strip() != ""].copy()

            rows.append(tmp[["연도", "구분", "관", "항", "목", "금액"]])

        except Exception:
            continue

    if not rows:
        return pd.DataFrame(columns=["연도", "구분", "관", "항", "목", "금액"])

    out = pd.concat(rows, ignore_index=True)
    out = out.groupby(["연도", "구분", "관", "항", "목"], as_index=False, sort=False)["금액"].sum()
    return out

# ======================================================
# 테마/색
# ======================================================
COMMON_FONT = dict(family="Arial", size=18)

def _theme_base() -> str:
    try:
        return (st.get_option("theme.base") or "").lower()
    except Exception:
        return ""

def _font_color() -> str:
    return "black" if _theme_base() == "light" else "white"

def _colors():
    base = _theme_base()
    if base == "light":
        return {"pos": "#1f77b4", "neg": "#d62728", "grid": "rgba(0,0,0,0.15)", "zero": "rgba(0,0,0,0.35)"}
    return {"pos": "#4da3ff", "neg": "#ff6b6b", "grid": "rgba(255,255,255,0.18)", "zero": "rgba(255,255,255,0.35)"}

def apply_common_layout(fig: go.Figure, height: int = 700):
    cols = _colors()
    base = _theme_base()

    if base == "light":
        paper_bg = "white"
        plot_bg = "white"
        font_color = "black"
    else:
        paper_bg = "rgba(0,0,0,0)"
        plot_bg = "rgba(0,0,0,0)"
        font_color = "white"

    fig.update_layout(
        height=height,
        font={**COMMON_FONT, "color": font_color},
        paper_bgcolor=paper_bg,
        plot_bgcolor=plot_bg,
        margin=dict(t=90, r=60, l=80, b=60),
    )
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor=cols["grid"], zeroline=True, zerolinecolor=cols["zero"])

# ======================================================
# 그래프
# ======================================================
def plot_recent_amount(recent: pd.DataFrame, title_label: str) -> go.Figure:
    fig = go.Figure()
    fc = _font_color()

    fig.add_trace(
        go.Bar(
            x=recent["연도_str"],
            y=recent["금액_백만원"],
            name="금액(백만원)",
            text=recent["금액_백만원"].map(lambda x: f"{x:,.0f} 백만원"),
            textposition="outside",
            textfont=dict(family="Arial", size=22, color="black"),
            hovertemplate="%{x}년<br>%{y:,.0f} 백만원<extra></extra>",
        )
    )

    fig.update_layout(
        height=800,
        margin=dict(t=140, r=60, l=60, b=60),

        title=dict(
            text=f"{title_label} | 최근 5개년",
            font=dict(family="Arial", size=24, color=fc),
            x=0.5,
            xanchor="center",
            y=0.98,
            yanchor="top",
        ),

        # ✅ X축
        xaxis=dict(
            type="category",
            title=dict(
                text="회계연도",
                font=dict(family="Arial", size=24)   # 🔥 여기로 이동
            ),
            tickfont=dict(family="Arial", size=24),
        ),

        # ✅ Y축
        yaxis=dict(
            title=dict(
                text="금액(백만원)",
                font=dict(family="Arial", size=24)   # 🔥 여기로 이동
            ),
            tickfont=dict(family="Arial", size=24),
            tickformat=",",
        ),

        showlegend=False,
    )
    apply_common_layout(fig)
    return fig

def plot_recent_pct(recent: pd.DataFrame) -> go.Figure:
    pct = recent.copy()
    cols = _colors()
    fc = _font_color()

    pct["pct_label"] = pct["증감률_%"].map(
        lambda x: "" if pd.isna(x) else f"{'▲' if x >= 0 else '▼'} {abs(x):.2f}%"
    )

    max_abs_pct = pd.to_numeric(pct["증감률_%"], errors="coerce").abs().max()
    if pd.isna(max_abs_pct):
        max_abs_pct = 0
    ylim = max(5, max_abs_pct * 1.3)

    fig = go.Figure()
    fig.add_bar(
        x=pct["연도_str"],
        y=pct["증감률_%"],
        text=pct["pct_label"],
        textposition="outside",
        textfont=dict(family="Arial", size=22, color="black"),
        marker_color=[cols["pos"] if (pd.notna(v) and v >= 0) else cols["neg"] for v in pct["증감률_%"]],
        hovertemplate="%{x}년<br>%{y:+.2f}%<extra></extra>",
    )
    fig.add_hline(y=0, line_color="gray", opacity=0.6)

    fig.update_layout(
        height=350,
        margin=dict(t=40, b=40, l=60, r=40),
        showlegend=False,

        xaxis=dict(
            type="category",
            title=dict(
                text="회계연도",
                font=dict(family="Arial", size=24)   # ✅ 여기
            ),
            tickfont=dict(family="Arial", size=24),
        ),

        yaxis=dict(
            title=dict(
                text="증감률(%)",
                font=dict(family="Arial", size=24)   # ✅ 여기
            ),
            tickfont=dict(family="Arial", size=24),
            ticksuffix="%",
            range=[-ylim, ylim],
        ),
    )

    apply_common_layout(fig)
    return fig

# ======================================================
# 표
# ======================================================
def render_table(series: pd.DataFrame):
    show_display = series.rename(
        columns={
            "금액": "금액(원)",
            "금액_백만원": "금액(백만원)",
            "증감_백만원": "증감(백만원)",
            "증감률_%": "증감률(%)",
        }
    ).copy()

    show_display["금액(원)"] = show_display["금액(원)"].map(lambda x: f"{x:,.0f}")
    show_display["금액(백만원)"] = show_display["금액(백만원)"].map(lambda x: f"{x:,.0f}")
    show_display["증감(백만원)"] = show_display["증감(백만원)"].map(lambda x: "" if pd.isna(x) else f"{x:,.0f}")
    show_display["증감률(%)"] = show_display["증감률(%)"].map(lambda x: "" if pd.isna(x) else f"{x:+.2f}%")

    st.dataframe(show_display[["연도", "금액(원)", "금액(백만원)", "증감(백만원)", "증감률(%)"]], use_container_width=True)

# ======================================================
# 렌더(메인)
# ======================================================
def render():
    st.subheader("📈 연도별 증감")

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
    
    def _section_open(title: str, desc: str = ""):
        st.markdown(
            f"""
            <div class="section-card">
            <div class="section-title">{title}</div>
            {"<div class='section-desc'>" + desc + "</div>" if desc else ""}
            """,
            unsafe_allow_html=True,
        )

    def _section_close():
        st.markdown("</div>", unsafe_allow_html=True)

    files = list_data_files()
    if not files:
        st.error("data/ 폴더에 엑셀 파일이 없습니다.")
        st.stop()

    # =========================
    # 제표 | 조회단위 | 구분 | 조회구분 (한 줄)
    # =========================
    c1, c2, c3, c4 = st.columns([1.2, 1.0, 1.2, 1.6])

    with c1:
        statement_type = st.radio(
            "제표",
            ["자금계산서", "재무상태표", "운영계산서"],
            horizontal=True,
            key="a1_stmt",
        )

    with c2:
        level = st.radio(
            "조회 단위",
            ["관", "항", "목"],
            horizontal=True,
            key="a1_level",
        )

    with c3:
        unit_type = st.radio(
            "구분",
            ["전체", "등록금", "비등록금"],
            horizontal=True,
            key="a1_unit",
        )

    with c4:
        if statement_type == "재무상태표":
            io_filter = st.radio(
                "조회 구분",
                ["전체", "자산", "부채/기본금"],
                horizontal=True,
                key="a1_io",
            )
        else:
            io_filter = st.radio(
                "조회 구분",
                ["전체", "수입", "지출"],
                horizontal=True,
                key="a1_io",
            )
    st.divider()
    # =========================
    # 데이터 구축
    # =========================
    value_col = "당기" if statement_type in ("재무상태표", "운영계산서") else "결산"

    ts = build_timeseries(statement_type, unit_type, value_col)
    if ts.empty:
        st.error("선택한 조건으로 모을 데이터가 없습니다. (파일/시트명 규칙 확인)")
        st.stop()

    if statement_type != "재무상태표" and io_filter in ("수입", "지출"):
        ts = ts[ts["구분"] == io_filter].copy()
        if ts.empty:
            st.warning(f"'{io_filter}'으로 필터링한 결과가 없습니다.")
            st.stop()

    ts = ts.copy()
    ts["관_norm"] = ts["관"].map(_norm)
    ts["항_norm"] = ts["항"].map(_norm)
    ts["목_norm"] = ts["목"].map(_norm)

    # =========================
    # 0원 제외용
    # =========================
    def _nonzero_norms(df: pd.DataFrame, col_norm: str) -> set[str]:
        s = df.groupby(col_norm, as_index=False)["금액"].sum()
        return set(s.loc[s["금액"].abs() > 0, col_norm])

    nonzero_guans = _nonzero_norms(ts, "관_norm")
    nonzero_hangs = _nonzero_norms(ts[ts["항_norm"].astype(str).str.strip() != ""], "항_norm")
    nonzero_moks = _nonzero_norms(ts, "목_norm")

    # =========================
    # 목 옵션
    # =========================
    EXCLUDE_MOKS = {
        _norm("유동자금"),
        _norm("기타유동자산"),
        _norm("예수금"),
        _norm("선수금"),
        _norm("기타유동부채"),
    }
    valid_mok_norms = {n for n in nonzero_moks if n and n not in EXCLUDE_MOKS}

    mok_order = get_mok_order_from_files(statement_type, unit_type)
    ordered_norms = [n for n in mok_order if n in valid_mok_norms]
    ordered_set = set(ordered_norms)

    rest_norms = []
    seen = set()
    for n in ts["목_norm"].tolist():
        if n in valid_mok_norms and n not in ordered_set and n not in seen:
            seen.add(n)
            rest_norms.append(n)

    final_mok_norms = ordered_norms + rest_norms

    norm_to_label = dict(
        ts.loc[ts["목_norm"].isin(valid_mok_norms), ["목_norm", "목"]]
        .drop_duplicates(subset=["목_norm"])
        .itertuples(index=False, name=None)
    )

    MOK_OPTIONS = [
        {"id": f"MOK__{n}", "label": norm_to_label.get(n, n), "kind": "direct_mok", "match_mok_norm": n}
        for n in final_mok_norms
    ]

    # =========================
    # 관/항 옵션
    # =========================
    def _ordered_unique(seq):
        out = []
        seen = set()
        for x in seq:
            s = str(x).strip()
            if not s or s in seen:
                continue
            seen.add(s)
            out.append(s)
        return out

    guan_order = get_guan_order_from_files(statement_type, unit_type)
    hang_order = get_hang_order_from_files(statement_type, unit_type)

    def build_guan_options(df: pd.DataFrame) -> list[dict]:
        guans_ts = _ordered_unique(df["관"].tolist())
        latest_set = set(guan_order)
        guans_latest = [g for g in guan_order if g in set(guans_ts)]
        rest = [g for g in guans_ts if g not in latest_set]
        guans = guans_latest + rest

        out = []
        for g in guans:
            gnorm = _norm(g)
            if gnorm not in nonzero_guans:
                continue

            if io_filter != "전체":
                if statement_type in ("자금계산서", "운영계산서"):
                    if IO_GUAN_GROUP.get(g, "기타") != io_filter:
                        continue
                elif statement_type == "재무상태표":
                    if BS_GUAN_GROUP.get(g, "기타") != io_filter:
                        continue

            out.append({"id": f"GUAN__{g}", "label": g, "kind": "direct_guan", "match_guan": g})
        return out

    def build_hang_options(df: pd.DataFrame) -> list[dict]:
        # ts에 존재하는 항(등장순서)
        hangs_ts = _ordered_unique(df["항"].tolist())
        hangs_ts_norm = {_norm(h) for h in hangs_ts if _norm(h)}

        # ✅ 최신 시트 순서 우선 + 나머지(등장순서)
        hangs_latest = [h for h in hang_order if _norm(h) in hangs_ts_norm]
        latest_norm_set = {_norm(h) for h in hangs_latest}

        rest = [h for h in hangs_ts if _norm(h) and _norm(h) not in latest_norm_set]
        hangs = hangs_latest + rest

        out = []
        for h in hangs:
            hn = _norm(h)
            if not hn:
                continue
            if hn not in nonzero_hangs:
                continue
            out.append({"id": f"HANG__{h}", "label": h, "kind": "direct_hang", "match_hang": h})
        return out

    GUAN_OPTIONS = build_guan_options(ts)
    # ✅ 관 단위 드롭다운 맨 아래에 "총계" 옵션 추가
    if statement_type == "자금계산서" and level == "관" and io_filter == "전체":
        GUAN_OPTIONS.append({
            "id": "GUAN__TOTAL_CASH_OUT",
            "label": "총 계",
            "kind": "cash_total_row",
        }) 
    # ✅ 재무상태표 관 단위에서 합성 관 추가
    if statement_type == "재무상태표" and level == "관":
        GUAN_OPTIONS.append({
            "id": "GUAN__ASSET_TOTAL",
            "label": "자산총계",
            "kind": "asset_total",
        })
        GUAN_OPTIONS.append({
            "id": "GUAN__LIABILITY_TOTAL",
            "label": "부채총계",
            "kind": "liability_total",
        })
    HANG_OPTIONS = build_hang_options(ts)

    # =========================
    # 옵션 id 맵 + 시계열 집계
    # =========================
    opt_by_id = {x["id"]: x for x in GUAN_OPTIONS}
    opt_by_id.update({x["id"]: x for x in HANG_OPTIONS})
    opt_by_id.update({x["id"]: x for x in MOK_OPTIONS})

    def series_from_direct_guan(match_guan: str) -> pd.DataFrame:
        gnorm = _norm(match_guan)
        special_norms = {_norm(x) for x in SPECIAL_GUAN_DIRECT}

        if gnorm in special_norms:
            sub = ts[
                (ts["관_norm"] == gnorm)
                & (ts["항_norm"] == "")
                & (ts["목_norm"] == gnorm)
            ].copy()
        else:
            sub = ts[ts["관_norm"] == gnorm].copy()

        if sub.empty:
            return pd.DataFrame(columns=["연도", "금액"])
        return sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")
    
    def series_asset_total() -> pd.DataFrame:
        sub = ts[ts["관_norm"].isin(ASSET_TOTAL_GUANS)].copy()
        if sub.empty:
            return pd.DataFrame(columns=["연도", "금액"])
        return sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")


    def series_liability_total() -> pd.DataFrame:
        sub = ts[ts["관_norm"].isin(LIABILITY_TOTAL_GUANS)].copy()
        if sub.empty:
            return pd.DataFrame(columns=["연도", "금액"])
        return sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")

    def series_from_direct_hang(match_hang: str) -> pd.DataFrame:
        hnorm = _norm(match_hang)
        sub = ts[ts["항_norm"] == hnorm].copy()
        if sub.empty:
            return pd.DataFrame(columns=["연도", "금액"])
        return sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")
    def series_total_cash_out() -> pd.DataFrame:
        # ✅ 자금계산서 지출 전체를 연도별 합산
        sub = ts[ts["구분"] == "지출"].copy()
        if sub.empty:
            return pd.DataFrame(columns=["연도", "금액"])
        return sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")

    _cache: Dict[str, pd.DataFrame] = {}

    def get_series(option_id: str) -> pd.DataFrame:
        if option_id in _cache:
            return _cache[option_id]

        o = opt_by_id[option_id]
        kind = o.get("kind")

        # ✅ 자금계산서 '자 금 지 출 총 계' 행 직접 추출
        if kind == "cash_total_row":
            res = series_cashsheet_row_total(unit_type, value_col)

        # ✅ 재무상태표 합성 관
        elif kind == "asset_total":
            res = series_asset_total()

        elif kind == "liability_total":
            res = series_liability_total()

        # ✅ 기존 로직
        elif kind == "direct_guan":
            res = series_from_direct_guan(o["match_guan"])

        elif kind == "direct_hang":
            res = series_from_direct_hang(o["match_hang"])

        elif kind == "direct_mok":
            sub = ts[ts["목_norm"] == o["match_mok_norm"]].copy()
            res = (
                sub.groupby("연도", as_index=False)["금액"].sum().sort_values("연도")
                if not sub.empty
                else pd.DataFrame(columns=["연도", "금액"])
            )

        else:
            res = pd.DataFrame(columns=["연도", "금액"])

        _cache[option_id] = res
        return res

    # =========================
    # ✅ 단일 선택박스 — 전체 폭 사용
    # =========================
    if level == "관":
        labels = [x["label"] for x in GUAN_OPTIONS if not x.get("hidden")]
        by_label = {x["label"]: x for x in GUAN_OPTIONS if not x.get("hidden")}
        box_label = "관 선택"
    elif level == "항":
        labels = [x["label"] for x in HANG_OPTIONS]
        by_label = {x["label"]: x for x in HANG_OPTIONS}
        box_label = "항 선택"
    else:
        labels = [x["label"] for x in MOK_OPTIONS]
        by_label = {x["label"]: x for x in MOK_OPTIONS}
        box_label = "목 선택"

    if not labels:
        st.warning("선택 가능한 항목이 없습니다. (필터 조건을 확인하세요)")
        st.stop()

    # ✅ (1) 항목선택 섹션 박스
    _section_open("🔎 항목 선택", "조회 단위에 맞는 항목을 선택하면 아래에서 요약→그래프→표로 이어집니다.")
    sel_label = st.selectbox(box_label, labels, key=f"a1_single_select_{level}")
    sel = by_label[sel_label]
    title_label = f"{sel_label} ({level})"
    st.caption(f"선택: **{title_label}**")
    _section_close()

    st.divider()

    # =========================
    # 이후 동일(그래프/표)
    # =========================
    series = get_series(sel["id"])
    if series.empty:
        st.warning("선택한 항목에 대해 합산 결과가 없습니다.")
        st.stop()

    series = series.copy()
    series["연도"] = series["연도"].astype(int)
    series = series.sort_values("연도").reset_index(drop=True)
    series["금액_백만원"] = series["금액"] / 1_000_000
    series["증감_백만원"] = series["금액_백만원"].diff()
    series["증감률_%"] = series["금액_백만원"].pct_change() * 100

    recent = series.tail(5).copy()
    recent["연도_str"] = recent["연도"].astype(str)

    # ✅ (3) 그래프 섹션 박스
    _section_open("📊 추이 그래프", "최근 5개년 금액과 증감률을 함께 봅니다.")

    st.markdown("### 🕔 최근 5개년 비교 (금액)")
    st.plotly_chart(plot_recent_amount(recent, title_label), use_container_width=True)

    st.markdown("### 📉 전년 대비 증감률")
    st.plotly_chart(plot_recent_pct(recent), use_container_width=True)

    _section_close()

    st.divider()

    # ✅ (4) 표 섹션 박스
    _section_open("📋 데이터 표", "연도별 금액/증감/증감률을 표로 확인합니다.")
    render_table(series)
    _section_close()
