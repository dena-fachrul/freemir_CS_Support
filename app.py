import streamlit as st
import pandas as pd
import re
import string
import io
import time
import xlsxwriter

# --- 1. CONFIGURATION ---
st.set_page_config(
    page_title="Freemir CS Support",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. CORE LOGIC FUNCTIONS ---

def generate_reason_code(n):
    """Membuat kode urut A, B, ... Z, AA, dst."""
    code = ""
    while n >= 0:
        code = string.ascii_uppercase[n % 26] + code
        n = (n // 26) - 1
    return code[::-1]

def clean_and_split_sku(sku_raw):
    """Membersihkan SKU: split +, enter, dan FR yang menempel."""
    if pd.isna(sku_raw):
        return []
    s = str(sku_raw)
    s = re.sub(r'[\+\,\n\r]', ' ', s) 
    s = re.sub(r'(FR)', r' \1', s)    
    tokens = [t.strip() for t in s.split()]
    return [t for t in tokens if t and t.startswith("FR")]

@st.cache_data(show_spinner=False)
def process_data(uploaded_file):
    """Fungsi utama pengolahan data."""
    output = io.BytesIO()

    try:
        df = pd.read_excel(uploaded_file, sheet_name="Detail", usecols="C,H,J")
    except ValueError:
        return None, None, None, "Error: Sheet 'Detail' tidak ditemukan. Cek file Excel Anda."
    except Exception as e:
        return None, None, None, f"Error membaca file: {e}"

    df.columns = ['Order_ID', 'Raw_SKU', 'Reason']
    
    # --- Step 1: Cleaning ---
    expanded_data = []
    for _, row in df.iterrows():
        order_id = row['Order_ID']
        raw_sku = row['Raw_SKU']
        reason = row['Reason'] if not pd.isna(row['Reason']) else "No Reason"
        
        skus_list = clean_and_split_sku(raw_sku)
        for sku in skus_list:
            expanded_data.append({
                'Order ID': order_id,
                'SKU': sku,
                'Reason': reason
            })

    df_raw = pd.DataFrame(expanded_data)
    if df_raw.empty:
        return None, None, None, "Tidak ada data SKU valid (awalan 'FR') ditemukan."

    # --- Step 2: Pivot & Stats ---
    pivot_df = pd.crosstab(df_raw['SKU'], df_raw['Reason'])
    pivot_df['Total Problems'] = pivot_df.sum(axis=1)
    pivot_df = pivot_df.sort_values(by='Total Problems', ascending=False)

    # --- Step 3: Codification ---
    calc_cols = ['Total Problems']
    reason_columns = [col for col in pivot_df.columns if col not in calc_cols]
    
    reason_map = {}
    code_map = {}
    
    for i, reason in enumerate(reason_columns):
        code = generate_reason_code(i)
        reason_map[reason] = code
        code_map[code] = reason

    total_by_sku_series = pivot_df[reason_columns].sum()
    total_by_order_series = df_raw.groupby('Reason')['Order ID'].nunique()

    # Apply Codes
    pivot_df = pivot_df.rename(columns=reason_map)
    df_summary = pivot_df.reset_index()

    sorted_codes = sorted(reason_map.values(), key=lambda x: (len(x), x))
    cols_order = ['SKU', 'Total Problems'] + sorted_codes
    cols_order = [c for c in cols_order if c in df_summary.columns]
    df_summary = df_summary[cols_order]

    # Calculate Totals Row
    numeric_cols = ['Total Problems'] + sorted_codes
    numeric_cols = [c for c in numeric_cols if c in df_summary.columns]
    total_row = df_summary[numeric_cols].sum().to_frame().T
    total_row['SKU'] = 'TOTAL'
    df_summary_final = pd.concat([df_summary, total_row], ignore_index=True)

    # Add Sequence Number
    df_summary_final.insert(0, 'No', range(1, len(df_summary_final) + 1))
    df_summary_final.iloc[-1, 0] = ''

    # Prepare Legend Data
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

    # --- Step 4: Excel Export ---
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Integrated Report')
        
        # Styles
        fmt_header = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFEB3B', 'text_wrap': True})
        fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_left   = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
        fmt_total  = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF59D'})

        col_raw, col_sum, col_leg = 0, len(df_raw.columns) + 2, len(df_raw.columns) + 2 + len(df_summary_final.columns) + 2

        # Write Tables
        # 1. Raw Data
        for i, col in enumerate(df_raw.columns): ws.write(0, col_raw + i, col, fmt_header)
        for r, row in df_raw.iterrows():
            ws.write(r+1, col_raw, row['Order ID'], fmt_center)
            ws.write(r+1, col_raw+1, row['SKU'], fmt_center)
            ws.write(r+1, col_raw+2, row['Reason'], fmt_left)
        ws.set_column(col_raw, col_raw+2, 20)

        # 2. Summary
        for i, col in enumerate(df_summary_final.columns): ws.write(0, col_sum + i, col, fmt_header)
        for r, row in df_summary_final.iterrows():
            fmt = fmt_total if r == len(df_summary_final)-1 else fmt_center
            for c, val in enumerate(row):
                ws.write(r+1, col_sum + c, val if pd.notna(val) else "", fmt)
        ws.set_column(col_sum+1, col_sum+1, 25)

        # 3. Legend
        for i, col in enumerate(df_legend.columns): ws.write(0, col_leg + i, col, fmt_header)
        for r, row in df_legend.iterrows():
            ws.write(r+1, col_leg, row['Code'], fmt_center)
            ws.write(r+1, col_leg+1, row['Cancel / Refund Detail'], fmt_left)
            ws.write(r+1, col_leg+2, row['Total by SKU'], fmt_center)
            ws.write(r+1, col_leg+3, row['Total by Order'], fmt_center)
        ws.set_column(col_leg+1, col_leg+1, 50)

    output.seek(0)
    
    # KPI Stats
    stats = {
        'total_orders': df_raw['Order ID'].nunique(),
        'total_issues': len(df_raw),
        'unique_skus': df_raw['SKU'].nunique()
    }
    
    return output, stats, df_legend, "Success"

# --- 3. UI LAYOUT ---

# Sidebar
with st.sidebar:
    st.header("Freemir CS Support")
    st.markdown("---")
    st.info("""
    **Panduan Penggunaan:**
    1.  Siapkan file Excel `.xlsx` / `.xls`.
    2.  Pastikan sheet bernama **"Detail"**.
    3.  Upload file di panel utama.
    4.  Klik tombol 'Mulai Analisa'.
    """)
    st.markdown("---")
    st.caption("v3.0 Final Stable • Freemir Data Team")

# Main Page
st.title("🛡️ Freemir Customer Service Support")
st.markdown("##### Automated SKU Issue Analyzer & Reporting Tool")
st.divider()

# File Uploader
uploaded_file = st.file_uploader("Upload File Laporan Order (Format Excel)", type=['xlsx', 'xls'])

if uploaded_file:
    # Tombol Action
    if st.button("🚀 Mulai Analisa Data", type="primary", use_container_width=True):
        
        # Interactive Status Container (New Feature)
        with st.status("Sedang memproses data...", expanded=True) as status:
            st.write("📂 Membaca file Excel...")
            time.sleep(0.5) 
            st.write("🧹 Membersihkan & Memisahkan SKU...")
            time.sleep(0.5)
            
            # Run Logic
            excel_data, stats, df_legend, msg = process_data(uploaded_file)
            
            if msg == "Success":
                st.write("📊 Menghitung statistik & pivot...")
                time.sleep(0.3)
                st.write("✅ Selesai!")
                status.update(label="Analisa Berhasil!", state="complete", expanded=False)
                
                st.divider()
                
                # --- DASHBOARD METRICS ---
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Unique Orders", f"{stats['total_orders']:,}")
                col2.metric("Total Issues Found", f"{stats['total_issues']:,}")
                col3.metric("Problematic SKUs", f"{stats['unique_skus']:,}")
                
                st.divider()

                # --- TABLE PREVIEW ---
                st.subheader("📋 Rincian Masalah (Live Preview)")
                st.markdown("Tabel berikut menunjukkan kode masalah dan jumlah kejadiannya secara real-time.")
                
                st.dataframe(
                    df_legend,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Code": st.column_config.TextColumn("Kode", width="small"),
                        "Cancel / Refund Detail": st.column_config.TextColumn("Detail Masalah"),
                        "Total by SKU": st.column_config.NumberColumn("Total Issues"),
                        "Total by Order": st.column_config.NumberColumn("Unique Orders"),
                    }
                )

                # --- DOWNLOAD BUTTON ---
                st.divider()
                st.success("Laporan siap diunduh!")
                
                col_dl1, col_dl2 = st.columns([1, 1])
                with col_dl1:
                    st.download_button(
                        label="📥 Download Laporan Excel",
                        data=excel_data,
                        file_name="Freemir_Integrated_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            else:
                status.update(label="Terjadi Kesalahan", state="error")
                st.error(msg)
