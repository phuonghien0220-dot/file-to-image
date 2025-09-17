import streamlit as st
import fitz  # PyMuPDF for PDF
from docx import Document
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
import os

st.set_page_config(page_title="Chuyển File sang Ảnh", layout="wide")

st.title("📄➡️🖼️ Tệp thành hình ảnh")

col1, col2 = st.columns([2,1])

with col1:
    uploaded_file = st.file_uploader(
        "Tải lên File", 
        type=["doc", "docx", "pdf", "xls", "xlsx"],
        help="Hỗ trợ: .doc, .docx, .pdf, .xls, .xlsx"
    )

with col2:
    st.subheader("⚙️ Tùy chọn chuyển đổi")

    file_type = None
    if uploaded_file:
        if uploaded_file.name.endswith((".doc", ".docx", ".pdf")):
            file_type = "word_pdf"
        elif uploaded_file.name.endswith((".xls", ".xlsx")):
            file_type = "excel"

    if file_type == "word_pdf":
        page_choice = st.radio("Chọn trang:", ["Tất cả trang", "Chọn khoảng trang"])
        if page_choice == "Chọn khoảng trang":
            page_range = st.text_input("Nhập khoảng trang (VD: 1-3,5)")

    if file_type == "excel":
        excel_option = st.radio("Chọn Sheet:", ["Tất cả", "Chọn một"])
        if excel_option == "Chọn một":
            sheet_name = st.text_input("Nhập tên sheet (VD: Sheet1)")
        cell_range = st.text_input("Nhập vùng dữ liệu (VD: A3:H20)", "")

    img_format = st.radio("Định dạng ảnh", ["PNG", "JPG", "WebP", "BMP"])
    dpi = st.slider("Chất lượng ảnh (DPI)", 72, 300, 150)

    convert_btn = st.button("🚀 Chuyển đổi")

# =============================
# Hàm xử lý (giả lập)
# =============================
if convert_btn and uploaded_file:
    st.success(f"✅ Đang xử lý file {uploaded_file.name} ...")

    if file_type == "word_pdf":
        st.info("👉 Hiện tại demo chỉ hỗ trợ PDF. Word sẽ cần chuyển sang PDF trước.")
        # Ví dụ xử lý PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            pdf_path = tmp.name
        pdf = fitz.open(pdf_path)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            pix = page.get_pixmap(dpi=dpi)
            img_path = f"page_{page_num+1}.{img_format.lower()}"
            pix.save(img_path)
            st.image(img_path, caption=f"Trang {page_num+1}")
            with open(img_path, "rb") as f:
                st.download_button(
                    f"Tải ảnh Trang {page_num+1}",
                    f,
                    file_name=img_path,
                    mime="image/"+img_format.lower()
                )

    elif file_type == "excel":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            excel_path = tmp.name
        xls = pd.ExcelFile(excel_path)

        if excel_option == "Tất cả":
            sheets = xls.sheet_names
        else:
            sheets = [sheet_name]

        for sh in sheets:
            df = pd.read_excel(excel_path, sheet_name=sh)
            if cell_range:
                df = df.loc[
                    df.index[int(cell_range[1:-2])-1: int(cell_range[-2:])],
                ]
            fig, ax = plt.subplots(figsize=(8,4))
            ax.axis('off')
            tbl = ax.table(cellText=df.values, colLabels=df.columns, loc='center')
            plt.tight_layout()
            img_path = f"{sh}.{img_format.lower()}"
            plt.savefig(img_path, dpi=dpi)
            st.image(img_path, caption=f"Sheet: {sh}")
            with open(img_path, "rb") as f:
                st.download_button(
                    f"Tải ảnh Sheet {sh}",
                    f,
                    file_name=img_path,
                    mime="image/"+img_format.lower()
                )
