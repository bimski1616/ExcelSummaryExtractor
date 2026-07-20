import pandas as pd

from openpyxl.utils.cell import (
    coordinate_from_string,
    column_index_from_string,
)

from config import MAPPING


def excel_cell(df, cell):

    col, row = coordinate_from_string(cell)

    row -= 1
    col = column_index_from_string(col) - 1

    return df.iloc[row, col]


def process_files(uploaded_files, progress_bar=None, status_text=None):

    hasil_semua = []
    error_log = []

    total_file = len(uploaded_files)

    for i, file in enumerate(uploaded_files, start=1):

        try:

            excel = pd.ExcelFile(file)

            # Cari sheet SUMMARY
            sheet_upper = [s.upper() for s in excel.sheet_names]

            if "SUMMARY" in sheet_upper:
                idx = sheet_upper.index("SUMMARY")
                sheet = excel.sheet_names[idx]
            else:
                sheet = excel.sheet_names[0]

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None
            )

            hasil = {
                "File Name": file.name,
                "Sheet Name": sheet
            }

            for field, cell in MAPPING.items():

                try:
                    hasil[field] = excel_cell(df, cell)

                except Exception:
                    hasil[field] = None

            hasil_semua.append(hasil)

        except Exception as e:

            error_log.append({
                "File Name": file.name,
                "Error": str(e)
            })

        # ======================
        # Progress
        # ======================

        if progress_bar is not None:
            progress_bar.progress(i / total_file)

        if status_text is not None:
            status_text.info(
                f"📄 Processing {i}/{total_file} : {file.name}"
            )

    hasil_df = pd.DataFrame(hasil_semua)

    error_df = pd.DataFrame(error_log)

    pd.set_option(
        "display.float_format",
        "{:,.2f}".format
    )

    return hasil_df, error_df