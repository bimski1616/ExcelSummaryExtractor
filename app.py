import streamlit as st
import pandas as pd
from io import BytesIO

from extractor import process_files

st.set_page_config(
    page_title="Excel Summary Extractor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Summary Extractor")

st.write(
    "Upload satu atau beberapa file Excel, lalu klik **Process**."
)

uploaded_files = st.file_uploader(
    "Upload Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.success(f"{len(uploaded_files)} file berhasil dipilih.")

    if st.button("🚀 Process"):

        # ==========================
        # Progress Bar
        # ==========================

        progress_bar = st.progress(0)

        status_text = st.empty()

        hasil_df, error_df = process_files(
            uploaded_files,
            progress_bar=progress_bar,
            status_text=status_text
        )

        progress_bar.empty()
        status_text.empty()

        st.success("✅ Semua file selesai diproses")

        # ==========================
        # Summary
        # ==========================

        total_file = len(uploaded_files)
        success_file = len(hasil_df)
        error_file = len(error_df)

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "📂 Total File",
            total_file
        )

        col2.metric(
            "✅ Success",
            success_file
        )

        col3.metric(
            "❌ Error",
            error_file
        )

        st.divider()

        # ==========================
        # PREVIEW
        # ==========================

        st.subheader("📊 Preview")

        preview_df = hasil_df.copy()

        # Format semua kolom numeric
        for col in preview_df.select_dtypes(include="number").columns:

            preview_df[col] = preview_df[col].map(
                lambda x: f"{x:,.2f}" if pd.notnull(x) else ""
            )

        st.dataframe(
            preview_df,
            use_container_width=True
        )

        st.divider()

        # ==========================
        # EXPORT
        # ==========================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            hasil_df.to_excel(
                writer,
                index=False,
                sheet_name="Summary"
            )

            worksheet = writer.sheets["Summary"]

            for idx, column in enumerate(hasil_df.columns, start=1):

                if pd.api.types.is_numeric_dtype(
                    hasil_df[column]
                ):

                    for row in range(
                        2,
                        worksheet.max_row + 1
                    ):

                        worksheet.cell(
                            row=row,
                            column=idx
                        ).number_format = "#,##0.00"

        output.seek(0)

        st.download_button(
            label="📥 Download Hasil Excel",
            data=output,
            file_name="hasil_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # ==========================
        # ERROR LOG
        # ==========================

        st.subheader("⚠ Error Log")

        if error_df.empty:

            st.success("Tidak ada file yang error.")

        else:

            st.error(
                f"{len(error_df)} file gagal diproses."
            )

            st.dataframe(
                error_df,
                use_container_width=True
            )
