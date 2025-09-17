import streamlit as st
import fitz  # PyMuPDF để xử lý PDF
from docx2pdf import convert as docx2pdf
from pdf2image import convert_from_path
import tempfile
import os
import zipfile
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from PIL import Image

st.set_page_config(page_title="Chuyển file sang ảnh", layout="wide")

st.title("📄➡️🖼️ Chuyển đổi file sang ảnh")

uploaded_file = st.file_uploader("Tải file Word (.docx, .doc), PDF, Excel (.xls, .xlsx)", 
                                 type=["docx", "doc", "pdf", "xls", "xlsx"])

def save_and_return_path(file):
    tmp_dir = tempfile.mkdtemp()
    file_path = os.path.join(tmp_dir, file.name)
    with open(file_path, "wb") as f:
        f.write(file.getbuffer())
    return file_path, tmp_dir

if uploaded_file:
    file_path, tmp_dir = save_and_return_path(uploaded_file)
    file_ext = uploaded_file.name.split(".")[-1].lower()

    # Xử lý PDF
    if file_ext == "pdf":
        doc = fitz.open(file_path)
        total_pages = len(doc)
        st.info(f"📑 File PDF có {total_pages} trang")

        selected_pages = st.multiselect(
            "Chọn trang muốn chuyển", list(range(1, total_pages+1)), default=[1]
        )

        if st.button("Chuyển sang ảnh"):
            img_list = []
            for p in selected_pages:
                page = doc.load_page(p-1)
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img_list.append(img)

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for i, img in enumerate(img_list, 1):
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PNG")
                    zipf.writestr(f"page_{i}.png", img_bytes.getvalue())
            st.download_button("⬇️ Tải ảnh (ZIP)", zip_buffer.getvalue(), "images.zip")

    # Xử lý Word
    elif file_ext in ["docx", "doc"]:
        # Chuyển Word sang PDF trước
        pdf_path = os.path.join(tmp_dir, "temp.pdf")
        docx2pdf(file_path, pdf_path)

        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        st.info(f"📑 File Word có {total_pages} trang (đã chuyển sang PDF)")

        selected_pages = st.multiselect(
            "Chọn trang muốn chuyển", list(range(1, total_pages+1)), default=[1]
        )

        if st.button("Chuyển sang ảnh"):
            img_list = []
            for p in selected_pages:
                page = doc.load_page(p-1)
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img_list.append(img)

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for i, img in enumerate(img_list, 1):
                    img_bytes = BytesIO()
                    img.save(img_bytes, format="PNG")
                    zipf.writestr(f"page_{i}.png", img_bytes.getvalue())
            st.download_button("⬇️ Tải ảnh (ZIP)", zip_buffer.getvalue(), "images.zip")

    # Xử lý Excel
    elif file_ext in ["xls", "xlsx"]:
        xls = pd.ExcelFile(file_path)
        sheet_name = st.selectbox("Chọn sheet", xls.sheet_names)
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        st.write("📊 Dữ liệu trong sheet")
        st.dataframe(df.head(10))

        cell_range = st.text_input("Nhập vùng dữ liệu (ví dụ A3:H20, để trống nếu muốn toàn bộ)", "")

        if st.button("Chuyển sang ảnh"):
            if cell_range:
                df_range = pd.read_excel(file_path, sheet_name=sheet_name, usecols=cell_range)
            else:
                df_range = df

            fig, ax = plt.subplots(figsize=(10, 5))
            ax.axis("off")
            tbl = ax.table(cellText=df_range.values, colLabels=df_range.columns, cellLoc="center", loc="center")
            tbl.auto_set_font_size(False)
            tbl.set_fontsize(10)
            tbl.scale(1.2, 1.2)

            img_buf = BytesIO()
            plt.savefig(img_buf, format="png", bbox_inches="tight")
            st.download_button("⬇️ Tải ảnh Excel", img_buf.getvalue(), "excel.png", mime="image/png")
