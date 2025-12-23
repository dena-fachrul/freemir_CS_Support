import streamlit as st
import pandas as pd
import re
import string
import io
import xlsxwriter

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="SKU Issue Analyzer", layout="wide")

st.title("📊 SKU Issue Auto-Analyzer")
st.markdown("""
This application automates the process of cleaning and analyzing SKU data from Order Reports.
It separates concatenated SKUs, calculates issue frequencies, and generates a codified report.

**How to use:**
1. Upload your Excel file (`.xlsx` or `.xls`).
2. Ensure the file contains a sheet named **"Detail"**.
3. Click **Process Data** and download the resulting report.
""")

# --- LOGIC FUNCTIONS ---

def generate_reason_code(n):
    """Generates codes like A, B, C... Z, AA, AB, etc."""
    code = ""
    while n >= 0:
        code = string.ascii_uppercase[n % 26] + code
        n = (n // 26) - 1
    return code[::-1]

def clean_and_split_sku(sku_raw):
    """Cleans and splits SKUs based on delimiters (+, newline) and 'FR' prefix."""
    if pd.isna(sku_raw):
        return []
    s = str(sku_raw)
    # Replace common delimiters with space
    s = re.sub(r'[\+\,\n\r]', ' ', s) 
    # Force space before 'FR' to handle concatenated cases like 'FR01FR02'
    s = re.sub(r'(FR)', r' \1', s)    
    tokens = [t.strip() for t in s.split()]
    # Filter only valid SKUs starting with 'FR'
    return [t for t in tokens if t and t.startswith("FR")]

def process_data(uploaded_file):
    # Create an in-memory buffer for the output file
    output = io.BytesIO()

    try:
        # Read specific columns: C (Order ID), H (Raw SKU), J (Reason)
        df = pd.read_excel(uploaded_file, sheet_name="Detail", usecols="C,H,J")
    except ValueError:
        return None, "Error: Sheet 'Detail' not found in the uploaded file."
    except Exception as e:
        return None, f"Error reading file: {e}"

    # Rename columns for internal processing
    df.columns = ['Order_ID', 'Raw_SKU', 'Reason']
    
    # --- DATA CLEANING ---
    expanded_data = []
    for _, row in df.iterrows():
        order_id = row['Order_ID']
        raw_sku = row['Raw_SKU']
        reason = row['Reason'] if not pd.isna(row['Reason']) else "No Reason"
        
        skus_list = clean_and_split_sku(raw_sku)
        
        # Explode logic: If multiple SKUs exist in one cell, create separate rows
        for sku in skus_list:
            expanded_data.append({
                'Order ID': order_id,
                'SKU': sku,
                'Reason': reason
            })

    # DataFrame 1: Raw Data Cleaned
    df_raw = pd.DataFrame(expanded_data)
    if df_raw.empty:
        return None, "No valid SKUs (starting with 'FR') were found in the data."

    # --- STATISTICS & ANALYSIS ---
    
    # 1. Pivot Table: SKU vs Reason
    pivot_df = pd.crosstab(df_raw['SKU'], df_raw['Reason'])
    
    # 2. Calculate Total Problems per SKU
    pivot_df['Total Problems'] = pivot_df.sum(axis=1)

    # 3. Sort by Total Problems (Descending)
    pivot_df = pivot_df.sort_values(by='Total Problems', ascending=False)

    # --- CODIFICATION & LEGEND PREPARATION ---
    
    # Identify reason columns (excluding calculated columns)
    calc_cols = ['Total Problems']
    reason_columns = [col for col in pivot_df.columns if col not in calc_cols]
    
    reason_map = {} # Maps Full Reason -> Code (A, B...)
    code_map = {}   # Maps Code -> Full Reason
    
    for i, reason in enumerate(reason_columns):
        code = generate_reason_code(i)
        reason_map[reason] = code
        code_map[code] = reason

    # Calculate Totals for Legend Table
    # A. Total by SKU: Sum of issues across all SKUs
    total_by_sku_series = pivot_df[reason_columns].sum()
    
    # B. Total by Order: Unique Order IDs affected by this reason
    total_by_order_series = df_raw.groupby('Reason')['Order ID'].nunique()

    # --- APPLY CODES TO SUMMARY TABLE ---
    pivot_df = pivot_df.rename(columns=reason_map)
    df_summary = pivot_df.reset_index()

    # Reorder Columns: SKU, Total Problems, A, B, C...
    sorted_codes = sorted(reason_map.values(), key=lambda x: (len(x), x))
    cols_order = ['SKU', 'Total Problems'] + sorted_codes
    # Filter to ensure columns exist
    cols_order = [c for c in cols_order if c in df_summary.columns]
    df_summary = df_summary[cols_order]

    # Calculate Grand Total Row for Summary
    numeric_cols = ['Total Problems'] + sorted_codes
    numeric_cols = [c for c in numeric_cols if c in df_summary.columns]
    
    total_row = df_summary[numeric_cols].sum().to_frame().T
    total_row['SKU'] = 'TOTAL'
    
    df_summary = pd.concat([df_summary, total_row], ignore_index=True)

    # Add 'No' column (Sequence number)
    df_summary.insert(0, 'No', range(1, len(df_summary) + 1))
    df_summary.iloc[-1, 0] = '' # Clear 'No' for the Total row

    # --- PREPARE LEGEND DATA ---
    legend_rows = []
    for code in sorted_codes:
        original_reason = code_map[code]
        legend_rows.append({
            'Code': code,
            'Cancel / Refund Detail': original_reason,
            'Total by SKU': total_by_sku_series.get(original_reason, 0),
            'Total by Order': total_by_order_series.get(original_reason, 0)
        })
    df_legend = pd.DataFrame(legend_rows)

    # --- EXPORT TO EXCEL ---
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Integrated Report')
        
        # Styles
        fmt_header = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFFF00', 'text_wrap': True})
        fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_left   = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
        fmt_total  = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFFF00'})
        
        # Table Positions (separated by 2 empty columns)
        col_raw = 0
        col_sum = len(df_raw.columns) + 2
        col_leg = col_sum + len(df_summary.columns) + 2

        # 1. WRITE RAW DATA TABLE
        for i, col in enumerate(df_raw.columns):
            ws.write(0, col_raw + i, col, fmt_header)
        for r, row in df_raw.iterrows():
            ws.write(r + 1, col_raw, row['Order ID'], fmt_center)
            ws.write(r + 1, col_raw + 1, row['SKU'], fmt_center)
            ws.write(r + 1, col_raw + 2, row['Reason'], fmt_left)
        
        ws.set_column(col_raw, col_raw, 20)      # Order ID
        ws.set_column(col_raw+1, col_raw+1, 25)  # SKU
        ws.set_column(col_raw+2, col_raw+2, 40)  # Reason

        # 2. WRITE SUMMARY TABLE
        for i, col in enumerate(df_summary.columns):
            ws.write(0, col_sum + i, col, fmt_header)
        for r, row in df_summary.iterrows():
            is_last = (r == len(df_summary) - 1)
            fmt = fmt_total if is_last else fmt_center
            
            for c, val in enumerate(row):
                v = val if pd.notna(val) else ""
                ws.write(r + 1, col_sum + c, v, fmt)
        
        ws.set_column(col_sum, col_sum, 5)        # No
        ws.set_column(col_sum+1, col_sum+1, 25)   # SKU
        ws.set_column(col_sum+2, col_sum+2, 15)   # Total Problem
        ws.set_column(col_sum+3, col_sum+len(df_summary.columns)-1, 5) # Codes

        # 3. WRITE LEGEND TABLE
        for i, col in enumerate(df_legend.columns):
            ws.write(0, col_leg + i, col, fmt_header)
        for r, row in df_legend.iterrows():
            ws.write(r + 1, col_leg, row['Code'], fmt_center)
            ws.write(r + 1, col_leg + 1, row['Cancel / Refund Detail'], fmt_left)
            ws.write(r + 1, col_leg + 2, row['Total by SKU'], fmt_center)
            ws.write(r + 1, col_leg + 3, row['Total by Order'], fmt_center)

        ws.set_column(col_leg, col_leg, 8)        # Code
        ws.set_column(col_leg+1, col_leg+1, 50)   # Detail
        ws.set_column(col_leg+2, col_leg+3, 15)   # Totals

    # Reset pointer to start of the file
    output.seek(0)
    return output, "Success"

# --- STREAMLIT UI ---
uploaded_file = st.file_uploader("Upload Excel File (Order-CS.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    if st.button("Process Data"):
        with st.spinner('Processing data, please wait...'):
            processed_data, status = process_data(uploaded_file)
            
            if status == "Success":
                st.success("Analysis Complete! Download your report below.")
                
                st.download_button(
                    label="📥 Download Integrated Report",
                    data=processed_data,
                    file_name="Integrated_SKU_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(status)