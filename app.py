import streamlit as st
import fitz  # PyMuPDF for PDF
import pandas as pd
import matplotlib.pyplot as plt
import tempfile, os, zipfile

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
        page_choice = st.radio("Chọn trang:", ["Tất cả", "Khoảng trang"])
        if page_choice == "Khoảng trang":
            page_range = st.text_input("Nhập khoảng trang (VD: 1-3,5)")

    if file_type == "excel":
        excel_option = st.radio("Chọn Sheet:", ["Tất cả", "Một sheet"])
        if excel_option == "Một sheet":
            sheet_name = st.text_input("Tên sheet (VD: Sheet1)")
        cell_range = st.text_input("Vùng dữ liệu (VD: A3:H20)", "")

    img_format = st.radio("Định dạng ảnh", ["PNG", "JPG", "WebP", "BMP"])
    dpi = st.slider("Chất lượng ảnh (DPI)", 72, 300, 150)

    convert_btn = st.button("🚀 Chuyển đổi")

# =============================
# Xử lý khi bấm nút
# =============================
if convert_btn and uploaded_file:
    st.success(f"✅ Đang xử lý file {uploaded_file.name} ...")
    temp_dir = tempfile.mkdtemp()
    output_files = []

    # ==== PDF ====
    if file_type == "word_pdf":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            pdf_path = tmp.name
        pdf = fitz.open(pdf_path)

        for page_num in range(len(pdf)):
            page = pdf[page_num]
            pix = page.get_pixmap(dpi=dpi)
            img_path = os.path.join(temp_dir, f"page_{page_num+1}.{img_format.lower()}")
            pix.save(img_path)
            output_files.append(img_path)

    # ==== Excel ====
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
            fig, ax = plt.subplots(figsize=(8,4))
            ax.axis('off')
            ax.table(cellText=df.values, colLabels=df.columns, loc='center')
            plt.tight_layout()
            img_path = os.path.join(temp_dir, f"{sh}.{img_format.lower()}")
            plt.savefig(img_path, dpi=dpi)
            output_files.append(img_path)

    # ==== Tải từng ảnh ====
    for f in output_files:
        with open(f, "rb") as file:
            st.download_button(
                label=f"Tải {os.path.basename(f)}",
                data=file,
                file_name=os.path.basename(f),
                mime="image/"+img_format.lower()
            )

    # ==== Tải tất cả (ZIP) ====
    zip_path = os.path.join(temp_dir, "all_images.zip")
    with zipfile.ZipFile(zip_path, "w") as zf:
        for f in output_files:
            zf.write(f, os.path.basename(f))
    with open(zip_path, "rb") as f:
        st.download_button(
            label="📦 Tải tất cả ảnh (ZIP)",
            data=f,
            file_name="all_images.zip",
            mime="application/zip"
        )
