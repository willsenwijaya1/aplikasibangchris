#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import math
from io import BytesIO

# Judul aplikasi
st.title("Pembagi File Excel Menjadi Beberapa Bagian")

# Upload file
uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

# Input jumlah baris per file
rows_per_file = st.number_input("Jumlah baris per file", min_value=1, value=300, step=50)

if uploaded_file is not None:
    try:
        # Baca file Excel
        df = pd.read_excel(uploaded_file, dtype=str)
        total_rows = df.shape[0]
        total_parts = math.ceil(total_rows / rows_per_file)

        st.write(f"Total baris dalam file: {total_rows}")
        st.write(f"Akan dibagi menjadi {total_parts} file")

        # Proses pembagian dan simpan ke dalam list
        excel_files = []
        for i in range(total_parts):
            start_row = i * rows_per_file
            end_row = start_row + rows_per_file
            df_part = df[start_row:end_row]

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_part.to_excel(writer, index=False, sheet_name=f"Part{i+1}")
            excel_files.append((f"file_part_{i+1}.xlsx", output.getvalue()))

        st.success("Pembagian file berhasil!")

        # Tampilkan tombol download untuk setiap file
        for filename, data in excel_files:
            st.download_button(
                label=f"Download {filename}",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca atau memproses file: {e}")

