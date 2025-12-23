import streamlit as st
import pandas as pd
import re
import string
import io
import xlsxwriter
import time

# --- 1. CONFIGURATION & DARK THEME CSS ---
st.set_page_config(
    page_title="freemir CS Support",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk Tampilan Dark Mode & Sidebar Gradient
st.markdown("""
    <style>
    /* Main Background - Dark Color */
    .stApp {
        background-color: #0E1117;
        color: #FAFAFA;
    }
    
    /* Sidebar Gradient Background */
    section[data-testid="stSidebar"] {
        background: rgb(2,0,36);
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
    }

    /* Metric Cards Styling (Dark Glassmorphism) */
    div[data-testid="metric-container"] {
        background-color: #1E293B;
        border: 1px solid #334155;
        padding: 15px;
        border-radius: 10px;
        color: white;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.5);
    }
    
    /* Button Styling */
    .stButton>button {
        width: 100%;
        background-color: #3B82F6; /* Blue freemir tone */
        color: white;
        border-radius: 8px;
        height: 3em;
        font-weight: 600;
        border: none;
    }
    .stButton>button:hover {
        background-color: #2563EB;
    }

    /* File Uploader Dark Style */
    div[data-testid="stFileUploader"] {
        background-color: #1E293B;
        padding: 20px;
        border-radius: 10px;
        border: 1px dashed #475569;
    }
    
    /* Table/Dataframe Header Color */
    thead tr th:first-child {display:none}
    tbody th {display:none}
    </style>
""", unsafe_allow_html=True)

# --- 2. LOGIC FUNCTIONS ---

def generate_reason_code(n):
    code = ""
    while n >= 0:
        code = string.ascii_uppercase[n % 26] + code
        n = (n // 26) - 1
    return code[::-1]

def clean_and_split_sku(sku_raw):
    if pd.isna(sku_raw):
        return []
    s = str(sku_raw)
    s = re.sub(r'[\+\,\n\r]', ' ', s) 
    s = re.sub(r'(FR)', r' \1', s)    
    tokens = [t.strip() for t in s.split()]
    return [t for t in tokens if t and t.startswith("FR")]

@st.cache_data(show_spinner=False)
def process_data(uploaded_file):
    output = io.BytesIO()

    try:
        df = pd.read_excel(uploaded_file, sheet_name="Detail", usecols="C,H,J")
    except ValueError:
        return None, None, None, "Error: Sheet 'Detail' not found. Please check your file."
    except Exception as e:
        return None, None, None, f"Error reading file: {e}"

    df.columns = ['Order_ID', 'Raw_SKU', 'Reason']
    
    # --- PROCESSING ---
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
        return None, None, None, "No valid SKUs (starting with 'FR') found."

    # --- PIVOT & STATS ---
    pivot_df = pd.crosstab(df_raw['SKU'], df_raw['Reason'])
    pivot_df['Total Problems'] = pivot_df.sum(axis=1)
    pivot_df = pivot_df.sort_values(by='Total Problems', ascending=False)

    # --- CODIFICATION ---
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

    df_summary_final.insert(0, 'No', range(1, len(df_summary_final) + 1))
    df_summary_final.iloc[-1, 0] = ''

    # --- PREPARE LEGEND FOR WEB DISPLAY & EXCEL ---
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

    # --- EXCEL WRITING ---
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet('Integrated Report')
        
        # Yellow Headers (Standard Excel)
        fmt_header = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFEB3B', 'text_wrap': True})
        fmt_center = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_left   = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
        fmt_total  = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#FFF59D'})

        col_raw, col_sum, col_leg = 0, len(df_raw.columns) + 2, len(df_raw.columns) + 2 + len(df_summary_final.columns) + 2

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
    
    stats = {
        'total_orders': df_raw['Order ID'].nunique(),
        'total_issues': len(df_raw),
        'total_skus': df_raw['SKU'].nunique()
    }
    
    return output, stats, df_legend, "Success"

# --- 3. SIDEBAR UI ---
with st.sidebar:
    st.title("freemir CS")
    st.markdown("### Customer Service Support")
    st.markdown("---")
    st.markdown("""
    **Panduan Penggunaan:**
    1. Siapkan file Excel (format `.xlsx`).
    2. Pastikan ada sheet bernama **"Detail"**.
    3. Upload file di sebelah kanan.
    4. Tunggu analisa selesai.
    """)
    st.caption("© 2024 Freemir Data Team")

# --- 4. MAIN UI ---
st.title("🛡️ freemir Customer Service Support")
st.markdown("### Automated SKU Issue Analyzer")
st.markdown("---")

uploaded_file = st.file_uploader("Upload File 'Order-CS.xlsx' disini", type=['xlsx', 'xls'])

if uploaded_file:
    with st.spinner('⚙️ Sedang menganalisa data...'):
        time.sleep(0.8) # Efek visual
        excel_data, stats, df_legend, status = process_data(uploaded_file)

    if status == "Success":
        st.success("✅ Analisa Selesai!")
        
        # --- METRICS ROW (3 Column) ---
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Total Unique Orders", stats['total_orders'])
        with c2: st.metric("Total Issues Found", stats['total_issues'])
        with c3: st.metric("Problematic SKUs", stats['total_skus'])

        # --- DOWNLOAD BUTTON ---
        st.markdown("### 📥 Download Report")
        col_dl_1, col_dl_2 = st.columns([3, 1])
        with col_dl_1:
            st.info("File output mencakup: Raw Data Cleaned, Summary Pivot, dan Legend Table.")
        with col_dl_2:
             st.download_button(
                label="Download Excel",
                data=excel_data,
                file_name=f"Freemir_SKU_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
             
        # --- TABLE DISPLAY (NEW REQUIREMENT) ---
        st.markdown("---")
        st.subheader("📋 Issue Summary & Legend Details")
        st.markdown("Berikut adalah rekapitulasi total masalah berdasarkan kategori (Data dari kolom Legend):")
        
        # Menampilkan Tabel df_legend di Web
        st.dataframe(
            df_legend, 
            use_container_width=True,
            hide_index=True,
            column_config={
                "Code": st.column_config.TextColumn("Kode", width="small"),
                "Cancel / Refund Detail": st.column_config.TextColumn("Detail Masalah", width="large"),
                "Total by SKU": st.column_config.NumberColumn("Total Issues", format="%d"),
                "Total by Order": st.column_config.NumberColumn("Total Orders", format="%d"),
            }
        )

    else:
        st.error(status)
else:
    st.info("👋 Silakan upload file Excel untuk memulai.")
