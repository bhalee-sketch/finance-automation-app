# pages/analysis2.py
# -*- coding: utf-8 -*-

from __future__ import annotations

import re
from typing import List, Dict, Tuple

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px  # ✅ 색 팔레트용

from statement.pages.utils import list_data_files, year_from_filename
from statement.pages.analysis1 import build_timeseries, apply_common_layout
from statement.pages.raw import parse_statement_sheets

# ✅ (선택) 도넛 클릭 이벤트용 - 설치되어 있으면 클릭 드릴다운, 없으면 selectbox 폴백
try:
    from streamlit_plotly_events import plotly_events
except Exception:
    plotly_events = None


# =========================
# ✅ 특수 관 구분 강제 매핑 (병하님 요청)
# - 계산 ❌ / 엑셀 값 그대로 ⭕
# - 단지 "구분"만 수입/지출로 강제
# =========================
SPECIAL_GUAN_IO_MAP = {
    "미사용전기이월자금": "수입",
    "미사용차기이월자금": "지출",
}

NO_DRILLDOWN_GUAN = {"미사용전기이월자금", "미사용차기이월자금"}
SPECIAL_GUAN_DIRECT = NO_DRILLDOWN_GUAN  # 의미를 분리하고 싶으면 따로 둬도 됨

def _norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s).replace("\u00a0", " ")).strip()


def _leading_spaces(text: str) -> int:
    """앞 공백 개수( NBSP 포함, 탭은 4칸 가정 )."""
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


def _available_years() -> list[int]:
    years: list[int] = []
    for p in list_data_files():
        ytxt = year_from_filename(p.stem)
        try:
            years.append(int(ytxt))
        except Exception:
            pass
    return sorted(set(years))


@st.cache_data(show_spinner=False)
def _latest_file_path_str() -> str | None:
    pairs = []
    for p in list_data_files():
        ytxt = year_from_filename(p.stem)
        try:
            y = int(ytxt)
        except Exception:
            continue
        pairs.append((y, str(p)))
    if not pairs:
        return None
    pairs.sort(key=lambda x: x[0], reverse=True)
    return pairs[0][1]

def _net_depr_in_mok_table(sub_h: pd.DataFrame) -> pd.DataFrame:
    """
    재무상태표 유형/무형고정자산 목 구성비용:
    - 감가상각누계액: 항상 음수로(차감) 반영
    - 사용수익권: '건물'에서 별도 차감(건물로 -합산), 사용수익권 조각은 제거
    """
    d = sub_h.copy()
    d["목"] = d["목"].fillna("").astype(str).str.replace("\u00a0", " ").str.strip()
    d["금액"] = pd.to_numeric(d["금액"], errors="coerce").fillna(0.0)

    # 1) 감가상각누계액 식별
    is_depr = d["목"].str.contains("감가상각누계액", na=False)

    # 2) 사용수익권 식별 (정확히 '사용수익권'이거나 포함하는 경우)
    is_use_right = d["목"].str.contains("사용수익권", na=False)

    # base(묶을 목 이름) 만들기
    base = d["목"].str.replace(r"\s+", "", regex=True)

    # 감가상각누계액은 base에서 제거(건물감가상각누계액 -> 건물)
    base = base.str.replace("감가상각누계액", "", regex=False)
    base = base.str.replace("누계액", "", regex=False)

    # ✅ 사용수익권은 '건물'로 강제 매핑(건물에서 빼야 하므로)
    base = base.where(~is_use_right, "건물")

    d["base_mok"] = base

    # 3) 부호 처리
    d["signed"] = d["금액"]

    # 감가상각누계액은 항상 음수로 (이미 -여도 -- 방지)
    d.loc[is_depr, "signed"] = -d.loc[is_depr, "금액"].abs()

    # 사용수익권도 '건물'에서 차감해야 하므로 항상 음수로
    d.loc[is_use_right, "signed"] = -d.loc[is_use_right, "금액"].abs()

    # 그 외 자산은 보통 양수(순액 구성비용)
    normal = ~(is_depr | is_use_right)
    d.loc[normal, "signed"] = d.loc[normal, "금액"].abs()

    # 4) base_mok 기준 순액 집계
    out = d.groupby("base_mok", as_index=False)["signed"].sum()
    out = out.rename(columns={"base_mok": "목", "signed": "금액"})

    # 5) 표시 정리: 빈값/0/음수는 도넛에서 제외(원하면 음수도 따로 표로 뽑을 수 있음)
    out["목"] = out["목"].astype(str).str.strip()
    out = out[(out["목"] != "")]
    out = out[pd.to_numeric(out["금액"], errors="coerce").fillna(0) > 0].copy()

    return out

def _theme_base() -> str:
    try:
        return (st.get_option("theme.base") or "").lower()
    except Exception:
        return ""

def _font_color() -> str:
    # 라이트면 검정, 다크면 흰색
    return "black" if _theme_base() == "light" else "white"

# =========================
# ✅ 도넛(라벨 밖/똑바로/순서 고정)
# - 색 이상함 방지: colorway 강제
# - undefined 제거: title_text=""
# =========================
def _plot_pie_outside(labels: list[str], values: list[float], height: int = 520) -> go.Figure:
    fig = go.Figure()

    fig.add_trace(
        go.Pie(
            labels=labels,
            values=values,
            hole=0.0,  # ✅ 도넛 ❌ → 원형 ⭕
            texttemplate=(
                "%{label}, %{customdata:,.0f} 백만원<br>"
                "%{percent:.0%}"
            ),
            customdata=[v / 1_000_000 for v in values],
            textposition="outside",
            automargin=True,
            sort=False,
            direction="clockwise",
            hovertemplate=(
                "%{label}<br>"
                "%{customdata:,.0f} 백만원<br>"
                "%{percent:.1%}<extra></extra>"
            ),
            marker=dict(
                line=dict(color="white", width=1.5)
            ),
        )
    )

    fig.update_layout(
        height=height,
        showlegend=False,
        margin=dict(t=20, b=20, l=40, r=40),
        font=dict(color="black", family="Arial", size=16),
        colorway=px.colors.qualitative.Plotly,
    )

    apply_common_layout(fig, height=height)

    # ✅ 최종 글씨색 고정 (theme 덮어쓰기 방지)
    fig.update_layout(font=dict(color="black"))
    fig.update_traces(textfont=dict(color="black"))

    return fig


@st.cache_data(show_spinner=False)
def _nested_orders_from_latest_sheet(
    statement_type: str, unit_type: str
) -> tuple[list[str], dict[str, list[str]], dict[tuple[str, str], list[str]]]:
    """
    최신 시트의 나열 순서대로:
      guan_order: [관...]
      hang_by_guan: {관: [항...]}
      mok_by_guan_hang: {(관,항): [목...]}
    """
    path_str = _latest_file_path_str()
    if not path_str:
        return [], {}, {}

    try:
        xls = pd.ExcelFile(path_str)
    except Exception:
        return [], {}, {}

    sheet_map = parse_statement_sheets(xls.sheet_names)
    sheet = sheet_map.get((statement_type, unit_type))
    if not sheet:
        return [], {}, {}

    try:
        df = pd.read_excel(path_str, sheet_name=sheet)
    except Exception:
        return [], {}, {}

    subj_candidates = ["과목", "계정", "항목", "과목명", "계정과목", "계정명"]
    subj = None
    for c in df.columns:
        if str(c).strip() in subj_candidates:
            subj = c
            break
    if subj is None:
        for c in df.columns:
            txt = str(c)
            if any(k in txt for k in subj_candidates):
                subj = c
                break
    if subj is None:
        return [], {}, {}

    guan_d, hang_d, mok_d = 0, 5, 10

    guan_order: list[str] = []
    hang_by_guan: dict[str, list[str]] = {}
    mok_by_guan_hang: dict[tuple[str, str], list[str]] = {}

    cur_g = ""
    cur_h = ""

    subjects = (
        df[subj].astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.rstrip()
        .tolist()
    )

    seen_g, seen_h, seen_m = set(), set(), set()

    for raw in subjects:
        if not str(raw).strip():
            continue

        depth = _leading_spaces(raw)
        name = str(raw).strip()

        if depth == guan_d:
            cur_g = name
            cur_h = ""
            k = _norm(cur_g)
            if k and k not in seen_g:
                seen_g.add(k)
                guan_order.append(cur_g)
            hang_by_guan.setdefault(cur_g, [])

        elif depth == hang_d:
            cur_h = name
            if cur_g:
                k = (_norm(cur_g), _norm(cur_h))
                if k not in seen_h:
                    seen_h.add(k)
                    hang_by_guan.setdefault(cur_g, []).append(cur_h)
                mok_by_guan_hang.setdefault((cur_g, cur_h), [])

        elif depth >= mok_d:
            if cur_g and cur_h:
                k = (_norm(cur_g), _norm(cur_h), _norm(name))
                if k not in seen_m:
                    seen_m.add(k)
                    mok_by_guan_hang.setdefault((cur_g, cur_h), []).append(name)

    return guan_order, hang_by_guan, mok_by_guan_hang


def _force_special_guan_io(df: pd.DataFrame) -> pd.DataFrame:
    """미사용전기이월자금=수입, 미사용차기이월자금=지출로 '구분'만 강제 보정 (계산 없음)"""
    if df.empty or "관" not in df.columns or "구분" not in df.columns:
        return df
    d = df.copy()
    d["관"] = d["관"].astype(str).str.strip()
    d["구분"] = d["구분"].astype(str).str.strip()

    mask = d["관"].isin(SPECIAL_GUAN_IO_MAP.keys())
    if mask.any():
        d.loc[mask, "구분"] = d.loc[mask, "관"].map(SPECIAL_GUAN_IO_MAP)
    return d


def render():
    st.subheader("📊 분석 2 | 드릴다운(관→항→목)")

    st.markdown(
        """
        <style>
        div[data-baseweb="select"] { cursor: pointer !important; }
        div[data-baseweb="select"] * { cursor: pointer !important; }
        <style>
        /* ✅ selectbox(우측 메뉴 포함) 커서: 손가락 */
        div[data-baseweb="select"] * { cursor: pointer !important; }
        div[data-baseweb="select"] input { cursor: pointer !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    years = _available_years()
    if not years:
        st.error("연도 파일을 찾지 못했습니다. (data 폴더 / 파일명 규칙 확인)")
        st.stop()

    # -------------------------
    # 상단 필터
    # -------------------------
    c1, c2, c3, c4 = st.columns([1.25, 1.10, 1.20, 1.0])

    with c1:
        statement_type = st.radio(
            "제표", ["자금계산서", "재무상태표", "운영계산서"],
            horizontal=False, key="a2_stmt"
        )
    with c2:
        unit_type = st.radio(
            "구분", ["전체", "등록금", "비등록금"],
            horizontal=False, key="a2_unit"
        )
    with c3:
        if statement_type == "재무상태표":
            io_filter = st.radio("조회구분", ["자산", "부채/기본금"], horizontal=False, key="a2_io")
        else:
            io_filter = st.radio("조회구분", ["수입", "지출"], horizontal=False, key="a2_io")
    with c4:
        year_sel = st.selectbox("회계연도", years, index=len(years) - 1, key="a2_year")

    top_level = st.radio(
        "상단 구성비 단위",
        ["관", "항", "목"],
        horizontal=True,
        key="a2_top_level",
    )
    # -------------------------
    # data load
    # -------------------------
    value_col = "당기" if statement_type in ("재무상태표", "운영계산서") else "결산"
    ts_all = build_timeseries(statement_type, unit_type, value_col)
    if ts_all.empty:
        st.error("선택한 조건으로 모을 데이터가 없습니다.")
        return

    ts_year_all = ts_all[ts_all["연도"] == int(year_sel)].copy()
    if ts_year_all.empty:
        st.warning("선택 연도의 데이터가 없습니다.")
        return

    # ✅ 미사용전기/차기이월자금: 계산 없이 값 그대로, 단 구분만 강제
    ts_year_all = _force_special_guan_io(ts_year_all)

    # ✅ 화면 도넛은 조회구분 필터된 데이터로
    ts_year = ts_year_all[ts_year_all["구분"] == io_filter].copy()
    if ts_year.empty:
        st.info(f"{io_filter} 데이터가 없습니다.")
        return

    # ✅ 최신 시트 기준 순서(관→항→목)
    guan_order, hang_by_guan, mok_by_guan_hang = _nested_orders_from_latest_sheet(statement_type, unit_type)

    # ==========================================================
    # ✅ (A) 드릴다운용: guan_tbl은 "항/목 드릴다운"에 필요하므로 항상 만든다
    # ==========================================================
    guan_tbl = ts_year.groupby("관", as_index=False)["금액"].sum()

    # ✅ 특수 관(미사용전기/차기)은 "관 헤더행(항 공백 + 목=관)" 값만 사용
    special_rows = ts_year.copy()
    for c in ["관", "항", "목"]:
        special_rows[c] = (
            special_rows[c]
            .fillna("")
            .astype(str)
            .str.replace("\u00a0", " ")
            .str.strip()
        )

    special_rows = special_rows[
        (special_rows["관"].isin(SPECIAL_GUAN_DIRECT)) &
        (special_rows["항"] == "") &
        (special_rows["목"].map(_norm) == special_rows["관"].map(_norm))
    ].copy()

    if not special_rows.empty:
        special_vals = special_rows.groupby("관")["금액"].sum().to_dict()
        guan_tbl["금액"] = guan_tbl.apply(
            lambda r: float(special_vals.get(str(r["관"]).strip(), r["금액"])),
            axis=1,
        )

    guan_tbl["관"] = guan_tbl["관"].astype(str).str.strip()
    guan_tbl = guan_tbl[(guan_tbl["관"] != "")]
    guan_tbl = guan_tbl[pd.to_numeric(guan_tbl["금액"], errors="coerce").fillna(0).abs() > 0].copy()

    if guan_tbl.empty:
        st.info("관 단위로 집계할 데이터가 없습니다.")
        return

    # ✅ 관 선택박스 순서(최신 시트 순서)
    exist_guans = set(guan_tbl["관"].unique())
    guan_order = [g for g in guan_order if g in exist_guans] or guan_tbl["관"].tolist()

    g_ord = {_norm(g): i for i, g in enumerate(guan_order)}
    guan_tbl["_ord"] = guan_tbl["관"].map(lambda x: g_ord.get(_norm(x), 10**9))
    guan_tbl = guan_tbl.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

    # ✅ 선택 관 초기화(드릴다운용)
    if "a2_sel_guan" not in st.session_state or st.session_state["a2_sel_guan"] not in set(guan_tbl["관"]):
        st.session_state["a2_sel_guan"] = str(guan_tbl.iloc[0]["관"])

    # ==========================================================
    # ✅ (B) 상단 구성비용: top_tbl은 top_level(관/항/목)에 따라 따로 만든다
    # ==========================================================
    top_col = {"관": "관", "항": "항", "목": "목"}[top_level]
    top_src = ts_year.copy()

    # ✅ 항/목으로 볼 때는 미사용전기/차기이월자금은 통째로 제외(혼선 방지)
    if top_level in ("항", "목"):
        top_src = top_src[~top_src["관"].astype(str).str.strip().isin(NO_DRILLDOWN_GUAN)].copy()

    top_tbl = top_src.groupby(top_col, as_index=False)["금액"].sum()
    top_tbl[top_col] = top_tbl[top_col].astype(str).str.strip()
    top_tbl = top_tbl[(top_tbl[top_col] != "")]
    top_tbl = top_tbl[pd.to_numeric(top_tbl["금액"], errors="coerce").fillna(0).abs() > 0].copy()

    if top_tbl.empty:
        st.info(f"{top_level} 단위로 집계할 데이터가 없습니다.")
        return

    exist_guans = set(guan_tbl["관"].unique())
    guan_order = [g for g in guan_order if g in exist_guans] or guan_tbl["관"].tolist()
    g_ord = {_norm(g): i for i, g in enumerate(guan_order)}
    guan_tbl["_ord"] = guan_tbl["관"].map(lambda x: g_ord.get(_norm(x), 10**9))
    guan_tbl = guan_tbl.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

    # ✅ 선택 관 초기화
    if "a2_sel_guan" not in st.session_state or st.session_state["a2_sel_guan"] not in set(guan_tbl["관"]):
        st.session_state["a2_sel_guan"] = str(guan_tbl.iloc[0]["관"])

    # ==========================================================
    # ✅ 1행: 관 구성비(전체폭)
    # ==========================================================
    st.markdown(f"### 🍩 {top_level} 구성비")
    fig_top = _plot_pie_outside(
        labels=top_tbl[top_col].astype(str).tolist(),
        values=top_tbl["금액"].astype(float).tolist(),
        height=550,
    )
    st.plotly_chart(fig_top, use_container_width=True)

    st.caption(f"선택 관: **{st.session_state['a2_sel_guan']}**")
    st.divider()

    # ==========================================================
    # ✅ 2행: 항 구성비 / 목 구성비 (2컬럼)
    # ==========================================================
    col_h, col_m = st.columns(2)

    # -------------------------
    # 항 구성비(왼쪽)
    # -------------------------
    with col_h:
        st.markdown("### 🍩 항 구성비")

        # ✅ 관 선택박스(도넛 위)
        sel_g = st.selectbox(
            "관 선택",
            guan_order,
            index=guan_order.index(st.session_state["a2_sel_guan"]) if st.session_state["a2_sel_guan"] in guan_order else 0,
            key="a2_sel_guan_box_under",
        )
        if sel_g != st.session_state["a2_sel_guan"]:
            st.session_state["a2_sel_guan"] = sel_g
            st.session_state.pop("a2_sel_hang", None)

        sel_g = st.session_state["a2_sel_guan"]

        # ✅ 드릴다운 제외 관이면: 항 도넛 계산/렌더 자체를 스킵
        if str(sel_g).strip() in NO_DRILLDOWN_GUAN:
            st.info("선택한 관은 하위(항/목) 구성비를 표시하지 않습니다.")
            st.session_state["a2_sel_hang"] = ""
            st.stop()  # ✅ col_h 블록 종료(성능 핵심)

        # ---- 여기부터는 드릴다운 가능한 관만 실행 ----
        sub_g = ts_year[ts_year["관"].astype(str).str.strip() == str(sel_g).strip()].copy()

        hang_tbl = sub_g.groupby("항", as_index=False)["금액"].sum()
        hang_tbl["항"] = hang_tbl["항"].astype(str).str.strip()
        hang_tbl = hang_tbl[(hang_tbl["항"] != "")]
        hang_tbl = hang_tbl[pd.to_numeric(hang_tbl["금액"], errors="coerce").fillna(0).abs() > 0].copy()

        if hang_tbl.empty:
            st.info("선택한 관 아래 항 데이터가 없습니다.")
            st.session_state["a2_sel_hang"] = ""
            st.stop()  # ✅ 항이 없으면 이후 렌더 스킵

        # ✅ 최신 시트 순서 적용
        hang_order = hang_by_guan.get(sel_g, [])
        exist_h = set(hang_tbl["항"].unique())
        hang_order = [h for h in hang_order if h in exist_h] or hang_tbl["항"].tolist()
        h_ord = {_norm(h): i for i, h in enumerate(hang_order)}
        hang_tbl["_ord"] = hang_tbl["항"].map(lambda x: h_ord.get(_norm(x), 10**9))
        hang_tbl = hang_tbl.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

        # ✅ 항 기본값
        if "a2_sel_hang" not in st.session_state or st.session_state["a2_sel_hang"] not in set(hang_tbl["항"]):
            st.session_state["a2_sel_hang"] = str(hang_tbl.iloc[0]["항"])

        fig_h = _plot_pie_outside(
            labels=hang_tbl["항"].astype(str).tolist(),
            values=hang_tbl["금액"].astype(float).tolist(),
            height=520,
        )

        # ✅ 여기서는 클릭 이벤트 없어도 됨(성능 우선) — 원하면 다시 plotly_events로 바꿀 수 있음
        st.plotly_chart(fig_h, use_container_width=True)


        # -------------------------
        # 목 구성비(오른쪽)
        # -------------------------
        with col_m:
            st.markdown("### 🍩 목 구성비")

            sel_g = st.session_state["a2_sel_guan"]
            sel_h = st.session_state.get("a2_sel_hang", "")

            # ✅ 드릴다운 제외 관이면: 목 도넛 스킵
            if str(sel_g).strip() in NO_DRILLDOWN_GUAN:
                st.info("선택한 관은 하위(항/목) 구성비를 표시하지 않습니다.")
                st.stop()

            # ✅ 항이 선택되지 않았으면: 목 도넛 계산/렌더 스킵 (성능 핵심)
            if not str(sel_h).strip():
                st.info("항을 선택하면 목 구성을 표시합니다.")
                st.stop()

            # ---- 여기부터는 (관+항) 선택이 있을 때만 실행 ----
            sub_g2 = ts_year[ts_year["관"].astype(str).str.strip() == str(sel_g).strip()].copy()

            hang_tbl2 = sub_g2.groupby("항", as_index=False)["금액"].sum()
            hang_tbl2["항"] = hang_tbl2["항"].astype(str).str.strip()
            hang_tbl2 = hang_tbl2[(hang_tbl2["항"] != "")]
            hang_tbl2 = hang_tbl2[pd.to_numeric(hang_tbl2["금액"], errors="coerce").fillna(0).abs() > 0].copy()

            if hang_tbl2.empty:
                st.info("선택한 관 아래 항이 없습니다.")
                st.stop()

            hang_order2 = hang_by_guan.get(sel_g, [])
            exist_h2 = set(hang_tbl2["항"].unique())
            hang_order2 = [h for h in hang_order2 if h in exist_h2] or hang_tbl2["항"].tolist()

            # ✅ 항 선택박스(도넛 위) — 이미 sel_h가 있지만, 최신 순서로 보정된 리스트를 보여주기 위함
            sel_h2 = st.selectbox(
                "항 선택",
                hang_order2,
                index=hang_order2.index(sel_h) if sel_h in hang_order2 else 0,
                key="a2_sel_hang_box_under",
            )
            st.session_state["a2_sel_hang"] = sel_h2
            sel_h = sel_h2

            sub_h = sub_g2[sub_g2["항"].astype(str).str.strip() == str(sel_h).strip()].copy()

            mok_tbl = sub_h.groupby("목", as_index=False)["금액"].sum()
            mok_tbl["목"] = mok_tbl["목"].astype(str).str.strip()
            mok_tbl = mok_tbl[(mok_tbl["목"] != "")]
            mok_tbl = mok_tbl[pd.to_numeric(mok_tbl["금액"], errors="coerce").fillna(0).abs() > 0].copy()

            if mok_tbl.empty:
                st.info("선택한 항 아래 목 데이터가 없습니다.")
                st.stop()

            mok_order = mok_by_guan_hang.get((sel_g, sel_h), [])
            exist_m = set(mok_tbl["목"].unique())
            mok_order = [m for m in mok_order if m in exist_m] or mok_tbl["목"].tolist()
            m_ord = {_norm(m): i for i, m in enumerate(mok_order)}
            mok_tbl["_ord"] = mok_tbl["목"].map(lambda x: m_ord.get(_norm(x), 10**9))
            mok_tbl = mok_tbl.sort_values("_ord").drop(columns=["_ord"]).reset_index(drop=True)

            fig_m = _plot_pie_outside(
                labels=mok_tbl["목"].astype(str).tolist(),
                values=mok_tbl["금액"].astype(float).tolist(),
                height=520,
            )
            st.plotly_chart(fig_m, use_container_width=True)

