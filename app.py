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
    convert_btn = st.button("🚀 Chuyển đổi")

def doc_to_pdf(doc_path, pdf_path):
    # Chuyển doc sang html rồi in ra pdf (cách đơn giản với Linux hoặc Streamlit Cloud)
    with open(doc_path, "rb") as doc_file:
        result = mammoth.convert_to_html(doc_file)
        html = result.value
    # Lưu html ra file tạm
    html_path = doc_path.replace(".doc", ".html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    # Dùng pandas hoặc thư viện khác để chuyển html sang pdf, hoặc dùng convert bằng tay nếu môi trường cho phép
    # Ở đây bạn nên chuyển file .doc sang .docx trước khi upload để đạt chất lượng tốt nhất

if convert_btn and uploaded_file:
    st.success(f"✅ Đang xử lý file {uploaded_file.name} ...")
    temp_dir = tempfile.mkdtemp()
    output_files = []

    file_name = uploaded_file.name.lower()
    file_ext = os.path.splitext(file_name)[1]

    # ==== Word (.docx, .doc) hoặc PDF ====
    if file_ext in [".docx", ".doc", ".pdf"]:
        # Nếu là Word .docx: chuyển sang PDF
        if file_ext == ".docx":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(uploaded_file.read())
                docx_path = tmp_docx.name
            pdf_path = os.path.join(temp_dir, "converted.pdf")
            docx2pdf_convert(docx_path, pdf_path)
        # Nếu là Word .doc: dùng mammoth chuyển sang html rồi pdf (khuyến nghị chuyển .doc sang .docx trước khi upload)
        elif file_ext == ".doc":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".doc") as tmp_doc:
                tmp_doc.write(uploaded_file.read())
                doc_path = tmp_doc.name
            # Chuyển .doc sang html
            with open(doc_path, "rb") as doc_file:
                result = mammoth.convert_to_html(doc_file)
                html = result.value
            # Lưu html ra file tạm
            html_path = os.path.join(temp_dir, "converted.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
            # Chuyển html sang pdf bằng pandas (hoặc bạn có thể render html ra ảnh trực tiếp)
            # Ở đây bạn nên chuyển file .doc sang .docx trước khi upload để đạt chất lượng tốt nhất!
            st.warning("Vui lòng chuyển file .doc sang .docx để đảm bảo chất lượng chuyển đổi tốt nhất.")
            pdf_path = None
        elif file_ext == ".pdf":
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                pdf_path = tmp_pdf.name

        # Đọc PDF và chuyển từng trang thành ảnh
        if pdf_path and os.path.exists(pdf_path):
            pdf = fitz.open(pdf_path)
            for page_num in range(len(pdf)):
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
        for sh in xls.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sh)
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.axis('off')
            ax.table(cellText=df.values, colLabels=df.columns, loc='center')
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
