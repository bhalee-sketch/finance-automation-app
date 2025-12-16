# xls_convert_app.py
# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO

import xlrd
from openpyxl import Workbook


def convert_xls_to_xlsx(uploaded_file) -> BytesIO:
    """
    업로드된 .xls 파일을 .xlsx 로 변환해서 BytesIO 로 반환.
    - xlrd 로 xls 읽고
    - openpyxl Workbook 으로 복사
    - 모든 시트, 모든 셀 값 그대로 복사(서식은 단순화)
    """
    # Streamlit UploadedFile -> bytes
    file_bytes = uploaded_file.read()

    # 1) xlrd로 .xls 열기
    book_xls = xlrd.open_workbook(file_contents=file_bytes, encoding_override="cp949")

    # 2) openpyxl 워크북 새로 생성
    wb_xlsx = Workbook()

    for sheet_idx in range(book_xls.nsheets):
        sheet_xls = book_xls.sheet_by_index(sheet_idx)

        # 첫 시트는 이미 있으니 제목만 바꾸고, 나머지는 새로 생성
        if sheet_idx == 0:
            ws = wb_xlsx.active
            ws.title = sheet_xls.name
        else:
            ws = wb_xlsx.create_sheet(title=sheet_xls.name)

        # 각 셀 값 복사
        for r in range(sheet_xls.nrows):
            row_values = sheet_xls.row_values(r)
            ws.append(row_values)

    # 3) 메모리로 저장해서 반환
    output = BytesIO()
    wb_xlsx.save(output)
    output.seek(0)
    return output


def run():
    # 상단 레이아웃: [뒤로가기 버튼] [제목 영역]
    back_col, title_col = st.columns([1, 5])

    with back_col:
        if st.button("← 메인으로"):
            # app.py 의 go("main")과 같은 역할
            st.session_state["page"] = "main"
            st.rerun()
            
    st.title("🔁 XLS → XLSX 변환")

    st.write("여러 개의 .xls 파일을 한 번에 업로드해서 각각 .xlsx로 변환합니다.")

    xls_files = st.file_uploader(
        "변환할 .xls 파일들을 선택하세요.",
        type=["xls"],
        accept_multiple_files=True,
    )

    # 🔹 따로 '변환' 버튼 없이, 업로드된 파일마다 바로 다운로드 버튼 생성
    if xls_files:
        st.info("각 파일 옆의 버튼을 눌러 .xlsx로 저장하세요.")

        for idx, xls_file in enumerate(xls_files):
            converted = convert_xls_to_xlsx(xls_file)

            base = xls_file.name.rsplit(".", 1)[0]
            out_name = base + ".xlsx"

            st.download_button(
                label=f"📥 {out_name} 다운로드",
                data=converted,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{idx}",
            )
