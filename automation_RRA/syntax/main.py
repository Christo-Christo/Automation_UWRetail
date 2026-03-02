import os
import pandas as pd
from syntax.RRA import read_data, process_final_data, generate_all_tables


# =============================
# Helper: nama file dari RAW DATA
# =============================
def build_output_name_from_raw(raw_data_path, suffix):
    name, ext = os.path.splitext(os.path.basename(raw_data_path))
    return f"{name}_{suffix}{ext}"


# =============================
# FILE 1 — DATA + PREVIOUS
# =============================
def write_data_file(final, previous, output_path, raw_data_path):
    output_file = os.path.join(
        output_path,
        build_output_name_from_raw(raw_data_path, "processed")
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final.to_excel(writer, sheet_name="Data", index=False)
        previous.to_excel(writer, sheet_name="Previous", index=False)

    print(f"✅ Data file saved: {output_file}")


# =============================
# FILE 2 — RC
# =============================
def write_rc_file(rc, rcsa_only, rc_sheets, output_path, raw_data_path):
    output_file = os.path.join(
        output_path,
        build_output_name_from_raw(raw_data_path, "RC")
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        rc.to_excel(writer, sheet_name="RC", index=False)
        rcsa_only.to_excel(writer, sheet_name="RCSA", index=False)

        for sheet_name, df in rc_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"✅ RC file saved: {output_file}")


# =============================
# FILE 3 — RB
# =============================
def write_rb_file(rb, rb_sheets, output_path, raw_data_path):
    output_file = os.path.join(
        output_path,
        build_output_name_from_raw(raw_data_path, "RB")
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        rb.to_excel(writer, sheet_name="RB", index=False)

        for sheet_name, df in rb_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

    print(f"✅ RB file saved: {output_file}")


# =============================
# MAIN (LOGIKA ASLI DIPERTAHANKAN)
# =============================
def main(input_path):

    # ---- BACA INPUT (ASLI, TIDAK DIUBAH) ----
    input_df = pd.read_excel(
        input_path,
        sheet_name="Input",
        engine="openpyxl",
        header=None
    )

    params = {}
    for _, row in input_df.iterrows():
        key = str(row[0]).strip().replace(":", "")
        params[key] = row[1]

    input_year = int(params["Year"])
    file_path_data = params["File Path Data"]
    file_path_output = params["File Path Output"]

    # ---- PIPELINE ASLI (TIDAK DIUBAH) ----
    data, previous, RCSA = read_data(file_path_data, input_path)
    final, col_name = process_final_data(data, previous, input_year)
    results = generate_all_tables(final, RCSA, input_year, col_name)

    # ---- PEMISAHAN HASIL (TANPA UBAH ISI) ----
    final_data = results["Data"]
    rc = results["RC"]
    rb = results["RB"]

    rc_sheets = {
        k: v for k, v in results.items()
        if "RCSA" in k or "NON-RCSA" in k
    }

    rb_sheets = {
        k: v for k, v in results.items()
        if k not in rc_sheets and k not in ["Data", "RC", "RB"]
    }

    rcsa_only = rc[rc["RCSA/NON-RCSA"] == "RCSA"]

    # ---- PENULISAN FILE (LOGIKA BARU, DATA SAMA) ----
    write_data_file(final_data, previous, file_path_output, file_path_data)
    write_rc_file(rc, rcsa_only, rc_sheets, file_path_output, file_path_data)
    write_rb_file(rb, rb_sheets, file_path_output, file_path_data)