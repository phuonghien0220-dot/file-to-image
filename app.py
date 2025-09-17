import streamlit as st
import tempfile
import os
import zipfile
from pdf2image import convert_from_path
from docx2pdf import convert as docx_to_pdf
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Convert File to Image", layout="centered")

st.title("📄➡️🖼️ Chuyển Word/PDF/Excel sang Ảnh")

uploaded_file = st.file_uploader("Tải lên file (.docx, .doc, .pdf, .xls, .xlsx)", 
                                 type=["docx", "doc", "pdf", "xls", "xlsx"])

if uploaded_file:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_ext}") as tmp:
        tmp.write(uploaded_file.read())
        input_path = tmp.name

    output_images = []

    # -------- PDF/Word --------
    if file_ext in ["pdf", "docx", "doc"]:
        if file_ext in ["docx", "doc"]:
            # Chuyển Word sang PDF trước
            pdf_path = input_path.replace(f".{file_ext}", ".pdf")
            docx_to_pdf(input_path, pdf_path)
        else:
            pdf_path = input_path

        st.info("Đang xử lý PDF/Word...")

        pages = convert_from_path(pdf_path, dpi=200)
        page_range = st.text_input("Nhập số trang (VD: 1-3 hoặc all)", "all")

        if st.button("Chuyển sang ảnh"):
            if page_range.lower() == "all":
                selected_pages = range(len(pages))
            else:
                a, b = [int(x) for x in page_range.split("-")]
                selected_pages = range(a-1, b)

            for i in selected_pages:
                out_file = f"page_{i+1}.png"
                pages[i].save(out_file, "PNG")
                output_images.append(out_file)

    # -------- Excel --------
    elif file_ext in ["xls", "xlsx"]:
        st.info("Đang xử lý Excel...")
        sheet_name = st.text_input("Tên sheet (để trống = sheet đầu tiên)")
        cell_range = st.text_input("Nhập vùng dữ liệu (VD: A1:H20, để trống = tất cả)")

        if st.button("Chuyển sang ảnh"):
            df = pd.read_excel(input_path, sheet_name=sheet_name if sheet_name else 0)

            if cell_range:
                import openpyxl
                wb = openpyxl.load_workbook(input_path, data_only=True)
                ws = wb[sheet_name if sheet_name else wb.sheetnames[0]]
                data = ws[cell_range]
                df = pd.DataFrame([[cell.value for cell in row] for row in data])

            out_file = "excel.png"
            dfi.export(df, out_file)
            output_images.append(out_file)

    # -------- Xuất kết quả --------
    if output_images:
        if len(output_images) == 1:
            with open(output_images[0], "rb") as f:
                st.download_button("⬇️ Tải ảnh", f, file_name=output_images[0])
        else:
            zip_name = "result.zip"
            with zipfile.ZipFile(zip_name, "w") as zf:
                for img in output_images:
                    zf.write(img)
            with open(zip_name, "rb") as f:
                st.download_button("⬇️ Tải tất cả ảnh (ZIP)", f, file_name=zip_name)

    # Xoá file tạm sau khi xong
    if os.path.exists(input_path):
        os.remove(input_path)
