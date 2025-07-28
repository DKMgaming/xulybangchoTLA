import streamlit as st
from docx import Document
import pandas as pd
import io

st.set_page_config(page_title="Trích xuất bảng từ Word", layout="wide")
st.title("📄 Trích xuất và làm phẳng bảng từ file Word (.docx)")

uploaded_file = st.file_uploader("Tải lên file Word (.docx)", type="docx")

if uploaded_file:
    # Đọc file Word
    doc = Document(uploaded_file)
    tables = doc.tables

    if not tables:
        st.warning("⚠️ Không tìm thấy bảng nào trong file.")
    else:
        st.success(f"✅ Đã tìm thấy {len(tables)} bảng trong file.")

        all_flattened_texts = []

        for idx, table in enumerate(tables):
            st.subheader(f"📊 Bảng {idx+1}")

            data = []
            for row in table.rows:
                data.append([cell.text.strip() for cell in row.cells])

            df = pd.DataFrame(data)
            st.dataframe(df)

            if df.shape[1] >= 3:
                st.markdown("### 🔄 Làm phẳng bảng (dạng văn bản)")
                flattened_texts = []
                for _, row in df.iloc[1:].iterrows():
                    freq, region3, vn = row[0], row[1], row[2]
                    text = f"Từ {freq}: Khu vực 3 sử dụng cho {region3}. Việt Nam sử dụng cho {vn}."
                    flattened_texts.append(text)

                all_flattened_texts.extend(flattened_texts)
                for t in flattened_texts:
                    st.write("- ", t)
            else:
                st.info("⚠️ Bảng này không đủ 3 cột để làm phẳng.")

        # Tải về file văn bản đã làm phẳng
        if all_flattened_texts:
            joined_text = "\n".join(all_flattened_texts)
            st.download_button("⬇️ Tải văn bản đã làm phẳng", joined_text.encode("utf-8"), file_name="flattened_texts.txt")
