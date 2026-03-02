import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from syntax.RRA import read_data, process_final_data, generate_all_tables


def read_input_sheet(input_path):
    """Membaca sheet Input untuk parameter"""
    input_df = pd.read_excel(input_path, sheet_name='Input', engine='openpyxl', header=None)
    
    params = {}
    for idx, row in input_df.iterrows():
        key = str(row[0]).strip().replace(":", "")  # Hapus hanya ":", jangan lowercase
        value = row[1]
        params[key] = value
    
    return params


def write_to_excel_with_format(results, output_path, input_path):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    import os

    input_filename = os.path.basename(input_path)
    name, ext = os.path.splitext(input_filename)
    output_file = os.path.join(output_path, f"{name}_processed{ext}")

    header_fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    def format_header(ws):
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font

    def format_grand_total(ws):
        for row in ws.iter_rows(min_row=2):
            if str(row[0].value).strip().lower() == "grand total":
                for cell in row:
                    cell.fill = header_fill
                    cell.font = header_font

    # --- Pisahkan sheet ---
    data_df = results.pop("Data")
    rc_df = results.pop("RC")
    rb_df = results.pop("RB")

    rc_sheets = {k: v for k, v in results.items() if "RCSA" in k or "NON-RCSA" in k}
    rb_sheets = {k: v for k, v in results.items() if k not in rc_sheets}

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # 1. Data
        data_df.to_excel(writer, sheet_name="Data", index=False)

        # 2. >>>
        pd.DataFrame().to_excel(writer, sheet_name=">>>", index=False)

        # 3. RC
        rc_df.to_excel(writer, sheet_name="RC", index=False)

        # 4. RC detail
        for name, df in rc_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

        # 5. >>>
        pd.DataFrame().to_excel(writer, sheet_name=">>>_RB", index=False)

        # 6. RB
        rb_df.to_excel(writer, sheet_name="RB", index=False)

        # 7. RB detail
        for name, df in rb_sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    # --- Formatting ---
    wb = load_workbook(output_file)

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        if ws.max_row >= 1:
            format_header(ws)
            format_grand_total(ws)

        # Auto width
        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    wb.save(output_file)
    print(f"✅ Output saved: {output_file}")


def main(input_path):
    """Main function"""
    print("📖 Reading input parameters...")
    params = read_input_sheet(input_path)
    
    input_year = int(params['Year'])
    file_path_data = params['File Path Data']
    file_path_output = params['File Path Output']
    
    print(f"📊 Year: {input_year}")
    print(f"📂 Data file: {file_path_data}")
    print(f"💾 Output folder: {file_path_output}")
    
    print("\n📥 Reading data files...")
    data, previous, RCSA = read_data(file_path_data, input_path)
    
    print("🔄 Processing final data...")
    final, col_name = process_final_data(data, previous, input_year)
    
    print("⚙️ Generating all summary tables (using ThreadPoolExecutor)...")
    results = generate_all_tables(final, RCSA, input_year, col_name)
    
    print("📝 Writing to Excel with formatting...")
    write_to_excel_with_format(results, file_path_output, input_path)
    
    print("✅ All done!")
