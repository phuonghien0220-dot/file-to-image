import streamlit as st
import fitz  # PyMuPDF
import tempfile, os, zipfile
from docx2pdf import convert as docx2pdf_convert

st.set_page_config(page_title="Chuyển Word sang Ảnh", layout="wide")
st.title("📄➡️🖼️ Chuyển Word thành Ảnh")

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "Tải lên file Word (.docx)", 
        type=["docx"],
        help="Chỉ hỗ trợ định dạng .docx"
    )

with col2:
    img_format = st.radio("Định dạng ảnh", ["PNG", "JPG"], horizontal=True)
    dpi = st.slider("Chất lượng ảnh (DPI)", 72, 300, 150)
    convert_btn = st.button("🚀 Chuyển đổi")

if convert_btn and uploaded_file:
    st.success(f"✅ Đang xử lý file {uploaded_file.name} ...")
    temp_dir = tempfile.mkdtemp()
    output_files = []

    # Lưu file Word tạm thời
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        tmp_docx.write(uploaded_file.read())
        docx_path = tmp_docx.name

    # Chuyển Word sang PDF
    pdf_path = os.path.join(temp_dir, "converted.pdf")
    docx2pdf_convert(docx_path, pdf_path)

    # Đọc PDF và chuyển từng trang thành ảnh
    pdf = fitz.open(pdf_path)
    for page_num in range(len(pdf)):
        page = pdf[page_num]
        pix = page.get_pixmap(dpi=dpi)
        img_path = os.path.join(temp_dir, f"page_{page_num+1}.{img_format.lower()}")
        pix.save(img_path)
        output_files.append(img_path)

    st.subheader("📥 Tải ảnh kết quả")
    for f in output_files:
        with open(f, "rb") as file:
            st.download_button(
                label=f"Tải {os.path.basename(f)}",
                data=file,
                file_name=os.path.basename(f),
                mime="image/"+img_format.lower()
            )

    # Tải tất cả ảnh dạng ZIP
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
