import streamlit as st
import fitz  # PyMuPDF để đọc PDF
from docx import Document
import pandas as pd
import matplotlib.pyplot as plt

st.title("📂 Chuyển đổi File sang Ảnh")

uploaded_file = st.file_uploader("Tải file (.docx, .pdf, .xlsx)", type=["docx", "pdf", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".pdf"):
        pdf = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        st.write(f"Tệp PDF có {len(pdf)} trang")
        page_numbers = st.multiselect("Chọn trang cần chuyển:", list(range(1, len(pdf)+1)))
        if st.button("Chuyển sang ảnh"):
            for p in page_numbers:
                page = pdf[p-1]
                pix = page.get_pixmap()
                st.image(pix.tobytes(), caption=f"Trang {p}")

    elif uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        st.write(f"Tệp Word có {len(doc.paragraphs)} đoạn văn")
        for i, para in enumerate(doc.paragraphs, 1):
            fig, ax = plt.subplots()
            ax.text(0.1, 0.5, para.text, fontsize=12)
            ax.axis("off")
            st.pyplot(fig)

    elif uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head())
        st.write("📸 Ảnh dữ liệu (5 dòng đầu)")
        fig, ax = plt.subplots()
        ax.axis("off")
        ax.table(cellText=df.head().values, colLabels=df.columns, loc="center")
        st.pyplot(fig)
