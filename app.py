import streamlit as st
from io import BytesIO
import tempfile
import zipfile
import os
import subprocess
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
import matplotlib.pyplot as plt
import string
import re
import shutil

# Try importing docx2pdf (optional, Windows/Mac)
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except Exception:
    DOCX2PDF_AVAILABLE = False

st.set_page_config(page_title="Converter → Images", layout="centered")

st.title("Chuyển file (Word / PDF / Excel) → ảnh")
st.write("Upload file .doc/.docx/.pdf/.xls/.xlsx. Chọn trang / vùng cần chuyển. Xuất ảnh .png (nhiều ảnh sẽ được nén .zip).")

# ---------- Helpers ----------
def col_letter_to_index(col):
    """A -> 0, B -> 1, AA -> 26"""
    col = col.upper()
    exp = 0
    total = 0
    for ch in col[::-1]:
        if ch in string.ascii_uppercase:
            total += (ord(ch) - ord('A') + 1) * (26**exp)
            exp += 1
    return total - 1

def excel_range_to_indices(rng):
    # rng like "A3:H20" or "B2"
    rng = rng.replace(" ", "")
    if ":" in rng:
        a,b = rng.split(":")
    else:
        a=b=rng
    match = re.match(r"([A-Za-z]+)(\d+)", a)
    match2 = re.match(r"([A-Za-z]+)(\d+)", b)
    if not match or not match2:
        raise ValueError("Vùng không hợp lệ. Ví dụ hợp lệ: A3:H20")
    col1, row1 = match.group(1), int(match.group(2))
    col2, row2 = match2.group(1), int(match2.group(2))
    r0 = row1 - 1
    r1 = row2 - 1
    c0 = col_letter_to_index(col1)
    c1 = col_letter_to_index(col2)
    # ensure ordering
    top, bottom = min(r0,r1), max(r0,r1)
    left, right = min(c0,c1), max(c0,c1)
    return top, bottom, left, right

def df_to_image(df, out_path, cell_width=120, cell_height=30, font_size=12):
    # Render dataframe to png using matplotlib table
    if df.shape[0] == 0:
        # create empty image
        img = Image.new("RGBA", (400, 200), (255,255,255,255))
        img.save(out_path)
        return
    nrows, ncols = df.shape
    # compute figure size
    fig_w = max(6, ncols * 1.2)
    fig_h = max(2, nrows * 0.4 + 1)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=120)
    ax.axis('off')
    # create table
    table = ax.table(cellText=df.values.astype(str),
                     colLabels=df.columns.astype(str) if df.columns is not None else None,
                     loc='center',
                     cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    table.scale(1, 1.2)
    plt.tight_layout()
    fig.savefig(out_path, bbox_inches='tight', pad_inches=0.3)
    plt.close(fig)

def word_to_pdf_via_libreoffice(input_path, output_dir):
    # use soffice headless to convert
    try:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir",
                        str(output_dir), str(input_path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # find pdf
        input_path = Path(input_path)
        pdf_path = Path(output_dir) / (input_path.stem + ".pdf")
        if pdf_path.exists():
            return str(pdf_path)
    except Exception as e:
        st.error(f"Lỗi khi chuyển Word->PDF bằng LibreOffice: {e}")
    return None

def docx_to_pdf(input_path, output_dir):
    # try docx2pdf first (works on Win/Mac)
    try:
        if DOCX2PDF_AVAILABLE:
            # docx2pdf can write into output dir if specified as second arg (only Windows?). Safer: convert then move.
            tmp_out = Path(output_dir) / (Path(input_path).stem + ".pdf")
            # docx2pdf convert may want absolute path
            docx2pdf_convert(str(input_path), str(tmp_out))
            if tmp_out.exists():
                return str(tmp_out)
    except Exception:
        pass
    # fallback to libreoffice
    return word_to_pdf_via_libreoffice(input_path, output_dir)

# ---------- UI ----------
uploaded = st.file_uploader("Chọn file", type=["pdf","doc","docx","xls","xlsx"], accept_multiple_files=False)

if uploaded is None:
    st.info("Chọn 1 file để bắt đầu.")
    st.stop()

# save uploaded to temp
tmpdir = tempfile.mkdtemp()
in_path = os.path.join(tmpdir, uploaded.name)
with open(in_path, "wb") as f:
    f.write(uploaded.getbuffer())

name, ext = os.path.splitext(uploaded.name)
ext = ext.lower()

output_images = []  # list of (filename, bytes)

if ext in [".pdf", ".doc", ".docx"]:
    st.subheader("Thiết lập chuyển trang (Word/PDF)")
    # For Word, convert to PDF first
    if ext in [".doc", ".docx"]:
        st.info("Chuyển Word -> PDF (sẽ dùng docx2pdf hoặc LibreOffice nếu có).")
        pdf_path = docx_to_pdf(in_path, tmpdir)
        if pdf_path is None:
            st.error("Không thể chuyển Word sang PDF tự động. Hãy cài LibreOffice (soffice) hoặc docx2pdf trên hệ thống.")
            st.stop()
    else:
        pdf_path = in_path

    # get number of pages
    try:
        pages = convert_from_path(pdf_path, first_page=1, last_page=1)
        # convert_from_path doesn't return number directly; use poppler's pdfinfo? but pdf2image has pdfinfo_from_path
        from pdf2image import pdfinfo_from_path
        info = pdfinfo_from_path(pdf_path)
        num_pages = info["Pages"]
    except Exception as e:
        st.error(f"Lỗi khi đọc file PDF: {e}")
        st.stop()

    st.write(f"Số trang trong file: **{num_pages}**")
    st.write("Chọn trang cần chuyển (ví dụ: 1,3-5 hoặc 'all'):")
    pages_sel = st.text_input("Trang (vd: all hoặc 1,3-5 hoặc 2-4)", value="all")
    out_format = st.selectbox("Định dạng ảnh:", ["png","jpg"], index=0)

    if st.button("Chuyển sang ảnh"):
        # parse pages_sel
        sel_pages = []
        if pages_sel.strip().lower() in ["all", "tất cả", "tat ca"]:
            sel_pages = list(range(1, num_pages+1))
        else:
            parts = pages_sel.split(",")
            for p in parts:
                p = p.strip()
                if "-" in p:
                    a,b = p.split("-")
                    sel_pages.extend(range(int(a), int(b)+1))
                else:
                    sel_pages.append(int(p))
            # clip
            sel_pages = [p for p in sel_pages if 1 <= p <= num_pages]
        if len(sel_pages) == 0:
            st.error("Không có trang hợp lệ được chọn.")
        else:
            with st.spinner("Đang chuyển..."):
                # convert selected pages
                images = convert_from_path(pdf_path, dpi=200, fmt=out_format, first_page=min(sel_pages), last_page=max(sel_pages))
                # convert_from_path returns pages in order from first_page..last_page; map to sel_pages carefully
                # Simpler: render individually
                for p in sel_pages:
                    try:
                        imgs = convert_from_path(pdf_path, dpi=200, fmt=out_format, first_page=p, last_page=p)
                        img = imgs[0]
                        bio = BytesIO()
                        img.save(bio, format=out_format.upper())
                        bio.seek(0)
                        filename = f"{name}_page{p}.{out_format}"
                        output_images.append((filename, bio.read()))
                    except Exception as e:
                        st.warning(f"Lỗi trang {p}: {e}")
            st.success(f"Hoàn tất: tạo được {len(output_images)} ảnh.")
            # if multiple images -> zip
            if len(output_images) > 1:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for fn, data in output_images:
                        zf.writestr(fn, data)
                zip_buffer.seek(0)
                st.download_button("Tải về .zip", data=zip_buffer, file_name=f"{name}_images.zip", mime="application/zip")
            else:
                fn, data = output_images[0]
                st.download_button("Tải ảnh", data=data, file_name=fn, mime=f"image/{out_format}")

elif ext in [".xls", ".xlsx"]:
    st.subheader("Thiết lập chuyển Excel -> ảnh")

    # read sheet names
    try:
        xl = pd.ExcelFile(in_path)
        sheets = xl.sheet_names
    except Exception as e:
        st.error(f"Lỗi đọc file Excel: {e}")
        st.stop()

    st.write("Sheet có trong file:")
    chosen_sheets = st.multiselect("Chọn sheet (chọn 1 hoặc nhiều hoặc để trống để chuyển tất cả):", options=sheets, default=[sheets[0]])
    # region input
    st.write("Bạn có thể nhập VÙNG (A1 style). Ví dụ: `A3:H20`. Nếu để trống, sẽ xuất toàn bộ sheet.")
    region = st.text_input("Vùng (ví dụ A3:H20) - để trống nghĩa là toàn bộ sheet", value="")
    out_format = st.selectbox("Định dạng ảnh:", ["png","jpg"], index=0)

    if st.button("Chuyển sang ảnh"):
        target_sheets = chosen_sheets if chosen_sheets else sheets
        images_all = []
        with st.spinner("Đang xử lý..."):
            for sh in target_sheets:
                try:
                    # read entire sheet (no header) so that row/col indices align with A1 coords
                    df_full = pd.read_excel(in_path, sheet_name=sh, header=None, engine=None)
                except Exception:
                    # fallback engine
                    df_full = pd.read_excel(in_path, sheet_name=sh, header=None)
                if region.strip() == "":
                    # use entire df but strip empty trailing rows/cols
                    df = df_full.fillna("")
                    # set column labels as A,B,C...
                    cols = []
                    for i in range(df.shape[1]):
                        # label by index (1-based)
                        cols.append(f"C{i+1}")
                    df.columns = cols
                else:
                    try:
                        top,bottom,left,right = excel_range_to_indices(region)
                        # slice df
                        df = df_full.iloc[top:bottom+1, left:right+1].fillna("")
                        # set column labels to original excel letters for clarity
                        col_labels = []
                        for c in range(left, right+1):
                            # convert index to letters
                            label = ""
                            n = c+1
                            while n > 0:
                                n, rem = divmod(n-1, 26)
                                label = chr(65+rem) + label
                            col_labels.append(label)
                        df.columns = col_labels
                    except Exception as e:
                        st.error(f"Lỗi vùng: {e}")
                        st.stop()

                # render df -> image file
                out_name = f"{name}_{sh}"
                img_path = os.path.join(tmpdir, f"{out_name}.png")
                try:
                    df_to_image(df, img_path)
                    with open(img_path, "rb") as f:
                        data = f.read()
                    images_all.append((f"{out_name}.png", data))
                except Exception as e:
                    st.warning(f"Lỗi khi tạo ảnh cho sheet {sh}: {e}")
            output_images = images_all

        if len(output_images) == 0:
            st.error("Không tạo được ảnh nào.")
        elif len(output_images) == 1:
            fn,data = output_images[0]
            st.download_button("Tải ảnh", data=data, file_name=fn, mime=f"image/png")
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fn,data in output_images:
                    zf.writestr(fn, data)
            zip_buffer.seek(0)
            st.download_button("Tải về .zip", data=zip_buffer, file_name=f"{name}_sheets_images.zip", mime="application/zip")

# cleanup temp dir on exit? (optional)
# shutil.rmtree(tmpdir)  # don't remove immediately to allow download
