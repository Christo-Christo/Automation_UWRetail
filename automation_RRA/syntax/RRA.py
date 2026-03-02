import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor, as_completed
import multiprocessing


def read_data(file_path_data, file_path_input):
    """Membaca semua data yang diperlukan"""
    data = pd.read_excel(file_path_data, sheet_name='sheet1', engine='openpyxl')
    previous = pd.read_excel(file_path_input, sheet_name='Previous', engine='openpyxl')
    RCSA = pd.read_excel(file_path_input, sheet_name='RCSA', engine='openpyxl')
    
    # Rename kolom previous
    previous = previous.rename(columns={'NoPolis': 'New Policy No.Check'})
    
    return data, previous, RCSA


def process_final_data(data, previous, input_year):
    """Proses data utama: merge, add columns, classifications"""
    
    # Merge dengan previous
    final = pd.merge(data, previous, on='ANO', how='left')
    
    # Pindahkan kolom New Policy No.Check ke posisi 3
    cols = final.columns.tolist()
    cols.insert(2, cols.pop(cols.index('New Policy No.Check')))
    final = final[cols]
    
    # Cek polis baru
    new_norm = final['NoPolis'].astype(str).str.strip()
    old_norm = final['New Policy No.Check'].astype(str).str.strip()
    final["Policy Status"] = np.where(
        old_norm.isna() | (old_norm == "nan") | (old_norm == ""),
        "New",
        np.where(new_norm.eq(old_norm), "Not new", "New")
    )
    cols = final.columns.tolist()
    cols.insert(3, cols.pop(cols.index("Policy Status")))
    final = final[cols]
    
    # Buat kolom-kolom baru
    final['TSI ZAI'] = final['AMOUNT TSI SHARE ADIRA']
    final['TSI Z'] = np.nan
    final['TZI IPZ'] = np.nan
    final['TREATY ZAI'] = final['QUOTASHARE-AMOUNT']
    final['TREATY Z'] = np.nan
    final['TREATY IPZ'] = np.nan
    final['BPPDAN AMOUNT'] = final['COMPULSORY-AMOUNT']
    
    # RB/RC
    final['RB/RC'] = np.where(
        final['TOC'].astype(str).str.contains(r'001|002|003|004|005|006|007|008|009', regex=True, na=False),
        'RB', 'RC'
    )
    
    # Kolom lainnya
    final['NAICS Code'] = np.nan
    final['Sustainability Checking Based on Insured Name'] = np.nan
    final['Sustainability Checking Based on NAICS Code'] = np.nan
    
    # Year classification
    col_name = f"Year {input_year}/Other Than {input_year}"
    final[col_name] = np.where(
        pd.to_datetime(final["StartDate"], errors="coerce").dt.year == int(input_year),
        str(input_year),
        f"Other than {input_year}"
    )
    cols = final.columns.tolist()
    cols.insert(13, cols.pop(cols.index(col_name)))
    final = final[cols]
    
    # Sharia/Conven
    final["Sharia/Conven"] = np.where(
        final["TOC"].astype(str).str.contains(r"Sharia|Syariah", case=False, regex=True, na=False),
        "Sharia", "Conven"
    )
    cols = final.columns.tolist()
    cols.insert(16, cols.pop(cols.index("Sharia/Conven")))
    final = final[cols]
    
    return final, col_name


def add_rcsa_classification(rc, RCSA):
    """Tambahkan kolom RCSA/NON-RCSA pada data RC"""
    rc_key = rc["RISK COORDINATE DESCRIPTION"].astype(str).str.strip()
    rcsa_key = RCSA["RISK COORDINATE"].astype(str).str.strip()
    
    rc_key_norm = rc_key.str.replace(" ", "").str.strip()
    rcsa_key_norm = rcsa_key.str.replace(" ", "").str.strip()
    
    rc.insert(
        17,
        "RCSA/NON-RCSA",
        rc_key_norm.isin(set(rcsa_key_norm)).map({True: "RCSA", False: "NON-RCSA"})
    )
    return rc


def create_summary_table(df, column_list, group_col, num_cols):
    """Membuat summary table dengan Grand Total"""
    result = (
        df[column_list]
        .groupby(group_col, as_index=False)
        .sum()
    )
    
    result[num_cols] = result[num_cols].apply(pd.to_numeric, errors="coerce")
    totals = result[num_cols].sum(numeric_only=True)
    
    grand_total_row = pd.DataFrame([{
        group_col: "Grand Total",
        "TREATY ZAI": totals["TREATY ZAI"],
        "SURPLUS 1-AMOUNT": totals["SURPLUS 1-AMOUNT"],
    }])
    
    result = pd.concat([result, grand_total_row], ignore_index=True)
    return result


def process_rc_category(rc, col_name, input_year, rcsa_status, year_filter, sharia_conven):
    """Process satu kategori RC"""
    column = ['RISK COORDINATE DESCRIPTION', 'TREATY ZAI', 'SURPLUS 1-AMOUNT']
    num_cols = ["TREATY ZAI", "SURPLUS 1-AMOUNT"]
    
    filtered = rc[
        (rc['RCSA/NON-RCSA'] == rcsa_status) &
        (rc[col_name].astype(str) == year_filter) &
        (rc['Sharia/Conven'] == sharia_conven)
    ]
    
    return create_summary_table(filtered, column, 'RISK COORDINATE DESCRIPTION', num_cols)


def process_rb_category(rb, col_name, year_filter, sharia_conven):
    """Process satu kategori RB"""
    column_rb = ['DESCRIPTION', 'TREATY ZAI', 'SURPLUS 1-AMOUNT']
    num_cols = ["TREATY ZAI", "SURPLUS 1-AMOUNT"]
    
    filtered = rb[
        (rb[col_name].astype(str) == year_filter) &
        (rb['Sharia/Conven'] == sharia_conven)
    ]
    
    return create_summary_table(filtered, column_rb, 'DESCRIPTION', num_cols)


def generate_all_tables(final, RCSA, input_year, col_name, max_workers=None):

    if max_workers is None:
        max_workers = max(2, multiprocessing.cpu_count() - 1)

    rc = final[final["RB/RC"].str.upper().eq("RC")].copy()
    rc = add_rcsa_classification(rc, RCSA)

    rb = final[final["RB/RC"].str.upper().eq("RB")].copy()

    rc_tasks = [
        (f"RCSA {input_year} Conven", process_rc_category, rc, col_name, input_year, "RCSA", str(input_year), "Conven"),
        (f"NON-RCSA {input_year} Conven", process_rc_category, rc, col_name, input_year, "NON-RCSA", str(input_year), "Conven"),
        (f"RCSA {input_year} Sharia", process_rc_category, rc, col_name, input_year, "RCSA", str(input_year), "Sharia"),
        (f"NON-RCSA {input_year} Sharia", process_rc_category, rc, col_name, input_year, "NON-RCSA", str(input_year), "Sharia"),
        (f"RCSA Other than {input_year} Conven", process_rc_category, rc, col_name, input_year, "RCSA", f"Other than {input_year}", "Conven"),
        (f"NON-RCSA Other than {input_year} Conven", process_rc_category, rc, col_name, input_year, "NON-RCSA", f"Other than {input_year}", "Conven"),
        (f"RCSA Other than {input_year} Sharia", process_rc_category, rc, col_name, input_year, "RCSA", f"Other than {input_year}", "Sharia"),
        (f"NON-RCSA Other than {input_year} Sharia", process_rc_category, rc, col_name, input_year, "NON-RCSA", f"Other than {input_year}", "Sharia"),
    ]

    rb_tasks = [
        (f"Conven {input_year}", process_rb_category, rb, col_name, str(input_year), "Conven"),
        (f"Sharia {input_year}", process_rb_category, rb, col_name, str(input_year), "Sharia"),
        (f"Conven Other than {input_year}", process_rb_category, rb, col_name, f"Other than {input_year}", "Conven"),
        (f"Sharia Other than {input_year}", process_rb_category, rb, col_name, f"Other than {input_year}", "Sharia"),
    ]

    results = {}
    future_to_name = {}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for name, func, *args in rc_tasks + rb_tasks:
            future = executor.submit(func, *args)
            future_to_name[future] = name

        for future in as_completed(future_to_name):
            sheet_name = future_to_name[future]
            results[sheet_name] = future.result()

    results["Data"] = final
    results["RC"] = rc
    results["RB"] = rb

    return results
