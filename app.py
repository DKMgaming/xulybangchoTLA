import streamlit as st
from docx import Document
import pandas as pd
import io
import fitz  # PyMuPDF
import pdfplumber

st.set_page_config(page_title="Trích xuất bảng từ Word và PDF", layout="wide")
st.title("📄 Trích xuất và làm phẳng bảng từ Word (.docx) hoặc PDF (.pdf)")

uploaded_file = st.file_uploader("Tải lên file Word (.docx) hoặc PDF (.pdf)", type=["docx", "pdf"])

all_flattened_texts = []

def flatten_table_row_with_headers(headers, row):
    parts = []
    for h, cell in zip(headers, row):
        h = str(h).strip()
        cell = str(cell).strip()
        if h and cell:
            parts.append(f"{h}: {cell}")
    return "; ".join(parts)

if uploaded_file:
    if uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        tables = doc.tables

        if not tables:
            st.warning("⚠️ Không tìm thấy bảng nào trong file Word.")
        else:
            st.success(f"✅ Đã tìm thấy {len(tables)} bảng trong file Word.")

            for idx, table in enumerate(tables):
                st.subheader(f"📊 Bảng {idx+1} (Word)")

                data = []
                for row in table.rows:
                    data.append([cell.text.strip() for cell in row.cells])

                df = pd.DataFrame(data)
                st.dataframe(df)

                if df.shape[0] > 1:
                    headers = df.iloc[0].tolist()
                    rows = df.iloc[1:].values.tolist()
                    st.markdown("### 🔄 Làm phẳng bảng (dạng văn bản có tiêu đề)")
                    flattened_texts = [flatten_table_row_with_headers(headers, row) for row in rows if any(row)]
                    all_flattened_texts.extend(flattened_texts)

                    for t in flattened_texts:
                        st.write("- ", t)

    elif uploaded_file.name.endswith(".pdf"):
        st.success("✅ Đang xử lý bảng trong file PDF...")
        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for tidx, table in enumerate(tables):
                    st.subheader(f"📄 Trang {page_num+1}, Bảng {tidx+1} (PDF)")
                    df = pd.DataFrame(table)
                    st.dataframe(df)

                    if df.shape[0] > 1:
                        headers = df.iloc[0].tolist()
                        rows = df.iloc[1:].values.tolist()
                        st.markdown("### 🔄 Làm phẳng bảng (dạng văn bản có tiêu đề)")
                        flattened_texts = [flatten_table_row_with_headers(headers, row) for row in rows if any(row)]
                        all_flattened_texts.extend(flattened_texts)

                        for t in flattened_texts:
                            st.write("- ", t)

# Tải về file văn bản đã làm phẳng
if all_flattened_texts:
    joined_text = "\n".join(all_flattened_texts)
    st.download_button("⬇️ Tải văn bản đã làm phẳng", joined_text.encode("utf-8"), file_name="flattened_texts.txt")
