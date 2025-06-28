import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
from datetime import datetime
import os

# 📁 저장 폴더 (서버 로컬 PC)
SAVE_FOLDER = "계획서기록"
os.makedirs(SAVE_FOLDER, exist_ok=True)

st.set_page_config(page_title="중량물 작업계획서", page_icon="📝", layout="centered")
st.title("📝 중량물 작업계획서 자동 생성기")

fields = {
    "부서명": "",
    "작업일자": datetime.today().strftime('%Y-%m-%d'),
    "작업 장소": "",
    "장비 No.": "",
    "직종": "",
    "성명": "",
    "작업 내용": "",
    "근무시간": "",
    "건강상태": "",
    "비고": ""
}

data = {}

with st.form("form"):
    for key, default in fields.items():
        data[key] = st.text_input(key, value=default)
    submitted = st.form_submit_button("📁 엑셀 파일 생성")

if submitted:
    wb = Workbook()
    ws = wb.active
    ws.title = "중량물 작업계획서"

    title_font = Font(size=14, bold=True)
    header_font = Font(size=12, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_wrap = Alignment(wrap_text=True)

    ws.merge_cells("A1:B1")
    ws["A1"] = "중량물 작업계획서"
    ws["A1"].font = title_font
    ws["A1"].alignment = align_center

    row = 3
    for key, value in data.items():
        ws[f"A{row}"] = key
        ws[f"B{row}"] = value
        ws[f"A{row}"].font = header_font
        ws[f"A{row}"].alignment = align_center
        ws[f"B{row}"].alignment = align_wrap
        row += 1

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 50

    # 📁 파일명 생성
    safe_name = data['부서명'].replace(" ", "_")
    filename = f"{data['작업일자']}_{safe_name}.xlsx"
    filepath = os.path.join(SAVE_FOLDER, filename)

    # 🖥️ 서버 로컬에 저장
    wb.save(filepath)

    # 🔽 다운로드용 버퍼 생성
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success(f"✅ 엑셀 파일이 생성되었습니다.\n\n📁 관리자 로컬에 저장됨: `{filepath}`")
    st.download_button(
        label="📥 브라우저로 엑셀 다운로드",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
