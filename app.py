import streamlit as st
import fitz  # PyMuPDF for PDF
import pandas as pd
import matplotlib.pyplot as plt
import tempfile, os, zipfile
from docx2pdf import convert as docx2pdf_convert
import mammoth

st.set_page_config(page_title="Chuyển File sang Ảnh", layout="wide")
st.title("📄➡️🖼️ Chuyển file thành ảnh (PNG/JPG)")

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "Tải lên file Word (.docx, .doc), PDF, Excel (.xls, .xlsx)", 
        type=["docx", "doc", "pdf", "xls", "xlsx"],
        help="Hỗ trợ: .docx, .doc, .pdf, .xls, .xlsx"
    )

with col2:
    img_format = st.radio("Định dạng ảnh", ["PNG", "JPG"], horizontal=True)
    dpi = st.slider("Chất lượng ảnh (DPI)", 72, 300, 150)
    file_type = None
    page_option = None
    page_range = ""
    sheet_option = None
    sheet_name = ""
    cell_range = ""

    if uploaded_file:
        file_name = uploaded_file.name.lower()
        file_ext = os.path.splitext(file_name)[1]
        if file_ext in [".docx", ".doc", ".pdf"]:
            file_type = "doc_pdf"
            page_option = st.radio("Chọn trang:", ["Tất cả", "Chọn trang cụ thể"])
            if page_option == "Chọn trang cụ thể":
                page_range = st.text_input("Nhập số trang (VD: 1-3,5)")
        elif file_ext in [".xls", ".xlsx"]:
            file_type = "excel"
            sheet_option = st.radio("Chọn sheet:", ["Tất cả", "Một sheet"])
            if sheet_option == "Một sheet":
                sheet_name = st.text_input("Tên sheet (VD: Sheet1)")
                cell_range = st.text_input("Vùng dữ liệu (VD: A3:H20)", "")
            else:
                cell_range = st.text_input("Vùng dữ liệu cho tất cả sheet (để trống nếu muốn tất cả)", "")

    convert_btn = st.button("🚀 Chuyển đổi")

def parse_page_range(page_range, total_pages):
    page_ids = []
    ranges = page_range.replace(" ", "").split(",")
    for r in ranges:
        if "-" in r:
            start, end = r.split("-")
            page_ids.extend(list(range(int(start)-1, int(end))))
        else:
            idx = int(r)-1
            if 0 <= idx < total_pages:
                page_ids.append(idx)
    # Loại bỏ trùng lặp
    page_ids = sorted(list(set([p for p in page_ids if 0 <= p < total_pages])))
    return page_ids

if convert_btn and uploaded_file:
    st.success(f"✅ Đang xử lý file {uploaded_file.name} ...")
    temp_dir = tempfile.mkdtemp()
    output_files = []

    file_name = uploaded_file.name.lower()
    file_ext = os.path.splitext(file_name)[1]

    # ==== Word (.docx, .doc) hoặc PDF ====
    if file_ext in [".docx", ".doc", ".pdf"]:
        # Chuyển Word (.docx) sang PDF
        if file_ext == ".docx":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(uploaded_file.read())
                docx_path = tmp_docx.name
            pdf_path = os.path.join(temp_dir, "converted.pdf")
            docx2pdf_convert(docx_path, pdf_path)
        # Chuyển Word (.doc) sang HTML, cảnh báo chất lượng chuyển đổi
        elif file_ext == ".doc":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as tmp_doc:
                tmp_doc.write(uploaded_file.read())
                doc_path = tmp_doc.name
            with open(doc_path, "rb") as doc_file:
                result = mammoth.convert_to_html(doc_file)
                html = result.value
            html_path = os.path.join(temp_dir, "converted.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
            st.warning("Khuyến nghị chuyển file .doc sang .docx để đảm bảo chất lượng tốt nhất.")
            pdf_path = None
        elif file_ext == ".pdf":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                pdf_path = tmp_pdf.name

        # Đọc PDF và chọn trang chuyển đổi
        if pdf_path and os.path.exists(pdf_path):
            pdf = fitz.open(pdf_path)
            total_pages = len(pdf)
            # Xác định trang cần chuyển
            if file_type == "doc_pdf" and page_option == "Chọn trang cụ thể" and page_range.strip():
                pages = parse_page_range(page_range, total_pages)
            else:
                pages = list(range(total_pages))
            for page_num in pages:
                page = pdf[page_num]
                pix = page.get_pixmap(dpi=dpi)
                img_path = os.path.join(temp_dir, f"page_{page_num+1}.{img_format.lower()}")
                pix.save(img_path)
                output_files.append(img_path)

    # ==== Excel (.xls, .xlsx) ====
    elif file_ext in [".xls", ".xlsx"]:
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp:
            tmp.write(uploaded_file.read())
            excel_path = tmp.name
        xls = pd.ExcelFile(excel_path)
        if sheet_option == "Một sheet" and sheet_name:
            sheets = [sheet_name] if sheet_name in xls.sheet_names else []
        else:
            sheets = xls.sheet_names

        for sh in sheets:
            df = pd.read_excel(excel_path, sheet_name=sh)
            df_show = df
            if cell_range:
                try:
                    # Chuyển vùng dữ liệu sang dạng slice
                    start, end = cell_range.replace(" ", "").split(":")
                    start_row = int(''.join(filter(str.isdigit, start))) - 1
                    start_col = ''.join(filter(str.isalpha, start))
                    end_row = int(''.join(filter(str.isdigit, end))) - 1
                    end_col = ''.join(filter(str.isalpha, end))
                    col_list = list(df.columns)
                    col_idx_start = col_list.index(start_col)
                    col_idx_end = col_list.index(end_col)
                    df_show = df.iloc[start_row:end_row+1, col_idx_start:col_idx_end+1]
                except Exception as e:
                    st.error(f"Vùng dữ liệu không hợp lệ: {e}")
                    df_show = df
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.axis('off')
            ax.table(cellText=df_show.values, colLabels=df_show.columns, loc='center')
            plt.tight_layout()
            img_path = os.path.join(temp_dir, f"{sh}.{img_format.lower()}")
            plt.savefig(img_path, dpi=dpi)
            plt.close(fig)
            output_files.append(img_path)

    # ==== Tải từng ảnh ====
    st.subheader("📥 Tải ảnh kết quả")
    for f in output_files:
        with open(f, "rb") as file:
            st.download_button(
                label=f"Tải {os.path.basename(f)}",
                data=file,
                file_name=os.path.basename(f),
                mime="image/"+img_format.lower()
            )

    # ==== Tải tất cả (ZIP) ====
    if output_files:
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
    else:
        st.error("Không có ảnh nào để tải về!")
