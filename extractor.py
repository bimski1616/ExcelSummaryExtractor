import pandas as pd

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from config import MAPPING
from project_rules import PROJECT_RULES


# =====================================================
# Ambil value berdasarkan alamat cell
# =====================================================
def get_cell(ws, cell):

    return ws[cell].value


# =====================================================
# Cari value berdasarkan label
#
# number_index
# 1 = angka pertama
# 2 = angka kedua
# =====================================================
def get_label_value(ws, label, number_index=1):

    label = label.strip().lower()

    for row in ws.iter_rows():

        for cell in row:

            if isinstance(cell, MergedCell):
                continue

            if cell.value is None:
                continue

            text = str(cell.value).strip().lower()

            if label not in text:
                continue

            start_row = cell.row
            start_col = cell.column

            # ============================================
            # PRIORITAS 1
            # Cari angka di kanan label
            # ============================================

            numbers = []

            for c in range(start_col + 1, start_col + 11):

                value = ws.cell(start_row, c).value

                if isinstance(value, (int, float)):
                    numbers.append(value)

            if len(numbers) >= number_index:
                return numbers[number_index - 1]

            # ============================================
            # PRIORITAS 2
            # Cari angka di bawah label
            # ============================================

            numbers = []

            for r in range(start_row + 1, start_row + 6):

                value = ws.cell(r, start_col).value

                if isinstance(value, (int, float)):
                    numbers.append(value)

            if len(numbers) >= number_index:
                return numbers[number_index - 1]

            # ============================================
            # PRIORITAS 3
            # Scan area sekitar
            # ============================================

            numbers = []

            for r in range(start_row, start_row + 5):

                for c in range(start_col, start_col + 10):

                    value = ws.cell(r, c).value

                    if isinstance(value, (int, float)):
                        numbers.append(value)

            if len(numbers) >= number_index:
                return numbers[number_index - 1]

    return None


# =====================================================
# Main Process
# =====================================================
def process_files(uploaded_files, progress_bar=None, status_text=None):

    hasil_semua = []
    error_log = []

    total_file = len(uploaded_files)

    for i, file in enumerate(uploaded_files, start=1):

        try:

            wb = load_workbook(
                file,
                data_only=True
            )

            # ============================================
            # Ambil sheet pertama
            # ============================================

            sheet = wb.sheetnames[0]
            ws = wb[sheet]

            # ============================================
            # Ambil Project Code
            # ============================================

            project_code = get_cell(ws, "J2")

            hasil = {
                "File Name": file.name,
                "Sheet Name": sheet,
                "PROJECT CODE": project_code
            }

            # ============================================
            # Loop semua mapping
            # ============================================

            for field, config in MAPPING.items():

                try:

                    if config["type"] == "cell":

                        hasil[field] = get_cell(
                            ws,
                            config["value"]
                        )

                    elif config["type"] == "label":

                        labels = config["value"]

                        if isinstance(labels, str):
                            labels = [labels]

                        # ====================================
                        # Default ambil angka pertama
                        # ====================================

                        number_index = (
                            PROJECT_RULES
                            .get(project_code, {})
                            .get(field, 1)
                        )

                        value = None

                        for lbl in labels:

                            value = get_label_value(
                                ws,
                                lbl,
                                number_index
                            )

                            if value is not None:
                                break

                        hasil[field] = value

                    else:

                        hasil[field] = None

                except Exception:

                    hasil[field] = None

            hasil_semua.append(hasil)

        except Exception as e:

            error_log.append({
                "File Name": file.name,
                "Error": str(e)
            })

        # ============================================
        # Progress Bar
        # ============================================

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
