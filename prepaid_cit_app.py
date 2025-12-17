# prepaid_cit_app.py
# -*- coding: utf-8 -*-

from io import BytesIO
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter


def run():
    # 상단: 뒤로가기 + 제목
    back_col, title_col = st.columns([1, 5])
    with back_col:
        if st.button("⬅ 메인으로"):
            st.session_state["page"] = "main"
            st.rerun()
    with title_col:
        st.title("🧾 선급법인세 취합")

    st.write("여러 엑셀을 업로드하면 선급법인세 자료를 제목행 포함 통합파일로 생성합니다.")

    uploaded = st.file_uploader(
        "엑셀 파일 업로드 (여러 개 가능)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("파일을 업로드하면 취합이 시작됩니다.")
        return

    # 가져올 열 (B, D, E, F, H, I, J, K, L)
    PICK_IDXS = [1, 3, 4, 5, 7, 8, 9, 10, 11]

    frames = []
    fail = []

    for f in uploaded:
        try:
            df = pd.read_excel(f, sheet_name=0, header=None)

            # 원본 1행 제거
            df = df.iloc[1:, :].dropna(how="all").reset_index(drop=True)

            max_col = df.shape[1]
            if any(i >= max_col for i in PICK_IDXS):
                fail.append((f.name, "필요한 열이 부족합니다"))
                continue

            sub = df.iloc[:, PICK_IDXS].copy()

            # ✅ 확장자 제거된 파일명만 사용
            filename = Path(f.name).stem
            sub.insert(0, "원본파일명", filename)

            frames.append(sub)

        except Exception as e:
            fail.append((f.name, str(e)))

    if not frames:
        st.error("취합할 데이터가 없습니다.")
        if fail:
            st.write(pd.DataFrame(fail, columns=["파일", "사유"]))
        return

    out = pd.concat(frames, ignore_index=True)

    # ✅ 제목행 정의
    out.columns = [
        "회계단위",
        "연월일",
        "예적금명",
        "예치기관",
        "사업자번호",
        "세율",
        "과세표준(수입이자)",
        "선급법인세",
        "법인지방소득세",
        "수입계정",
    ]

    st.success(f"취합 완료: {len(out):,}행")
    st.dataframe(out, use_container_width=True)

    # 엑셀 저장
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(
            writer,
            index=False,
            header=True,
            sheet_name="선급법인세_통합",
        )

        ws = writer.book["선급법인세_통합"]

        # 열 너비
        widths = [13, 10, 47, 13, 15.88, 4.6, 18, 22.13, 14.5, 17.25]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # (보너스) 제목행 고정
        ws.freeze_panes = "A2"

    buf.seek(0)

    st.download_button(
        "📥 통합파일 다운로드 (XLSX)",
        data=buf,
        file_name="선급법인세_통합.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
