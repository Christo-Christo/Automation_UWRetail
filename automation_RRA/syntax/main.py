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


def write_to_excel_with_format(results, output_path, input_year):
    """Menulis semua DataFrame ke Excel dengan formatting"""
    
    output_file = os.path.join(output_path, f"RRA_Output_{input_year}.xlsx")
    
    # Tulis semua DataFrame ke Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Format header dengan openpyxl
    wb = load_workbook(output_file)
    
    # Style untuk header
    header_fill = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")  # Biru terang
    header_font = Font(bold=True, color="FFFFFF")  # Putih bold
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Apply formatting ke baris pertama (header)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
        
        # Auto-fit column width (opsional)
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    # Tab colors
    tab_colors_2025 = ["RCSA 2025 Conven", "NON-RCSA 2025 Conven", "RCSA 2025 Sharia", "NON-RCSA 2025 Sharia", "Conven 2025", "Sharia 2025"]
    tab_colors_other = ["RCSA Other than 2025 Conven", "NON-RCSA Other than 2025 Conven", "RCSA Other than 2025 Sharia", "NON-RCSA Other than 2025 Sharia", "Conven Other than 2025", "Sharia Other than 2025"]
    
    for sheet_name in tab_colors_2025:
        if sheet_name in wb.sheetnames:
            wb[sheet_name].sheet_properties.tabColor = "C0E6F5"
    
    for sheet_name in tab_colors_other:
        if sheet_name in wb.sheetnames:
            wb[sheet_name].sheet_properties.tabColor = "DAF2D0"
    
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
    write_to_excel_with_format(results, file_path_output, input_year)
    
    print("✅ All done!")
