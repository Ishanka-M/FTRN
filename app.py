import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import io
import time

# --- CONFIGURATION ---
SL_TIMEZONE = pytz.timezone('Asia/Colombo')
GOOGLE_SHEET_NAME = "FTRN" 

AX_COLUMNS_TO_KEEP = [
    "Sales Order", "Style No", "FTR No", "Req. Date", "Issued By", 
    "Issued Date", "Item", "Color", "Shade No", "Pattern No", 
    "Serial No", "Batch number", "Unit", "Issue Qty"
]

# --- PAGE SETUP & BRANDING ---
st.set_page_config(page_title="FTRN Inventory System", layout="wide", page_icon="📦")

# CSS Animations & Branding
st.markdown("""
    <style>
    @keyframes slideIn {
        0% { transform: translateY(-20px); opacity: 0; }
        100% { transform: translateY(0); opacity: 1; }
    }
    .main-title {
        animation: slideIn 1s ease-out;
        text-align: center;
        color: #2E86C1;
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: transparent;
        color: #888888;
        text-align: center;
        padding: 10px;
        font-size: 14px;
        font-weight: bold;
    }
    </style>
    <h1 class='main-title'>📦 FTRN Inventory Control System</h1>
""", unsafe_allow_html=True)


# --- GOOGLE SHEETS CONNECTION ---
@st.cache_resource
def get_gspread_client():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

def get_sheet():
    client = get_gspread_client()
    return client.open(GOOGLE_SHEET_NAME).sheet1

# --- ADMIN AUTHENTICATION ---
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False

with st.sidebar:
    st.header("🔐 Admin Panel")
    if not st.session_state.admin_logged_in:
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            if username == "admin" and password == "admin@123":
                st.session_state.admin_logged_in = True
                st.toast("Admin ලොග් වීම සාර්ථකයි!", icon="✅")
                time.sleep(1)
                st.rerun()
            else:
                st.error("❌ Username හෝ Password වැරදියි!")
    else:
        st.success("Admin Logged In")
        if st.button("Logout"):
            st.session_state.admin_logged_in = False
            st.rerun()
        
        st.divider()
        st.subheader("🛠️ Manage Google Sheet")
        if st.button("⚠️ Reset Entire Sheet"):
            ws = get_sheet()
            ws.clear()
            st.toast("Sheet එක සම්පූර්ණයෙන්ම මකා දමන ලදී!", icon="🗑️")
            st.snow()

# --- MAIN APP LOGIC ---
st.divider()
col_remark, col_upload = st.columns(2)
with col_remark:
    remark = st.text_input("📝 Remark එක ඇතුළත් කරන්න:")
with col_upload:
    uploaded_file = st.file_uploader("📂 Excel File එක Upload කරන්න", type=["xlsx"])

if uploaded_file and remark:
    # 1. Load Data
    try:
        df_ax_raw = pd.read_excel(uploaded_file, sheet_name='AX')
        df_re = pd.read_excel(uploaded_file, sheet_name='RE')
    except Exception as e:
        st.error("Error: 'AX' සහ 'RE' sheet දෙකම file එකේ තිබිය යුතුය.")
        st.stop()

    # 2. Process AX Sheet
    df_ax = df_ax_raw.loc[:, ~df_ax_raw.columns.str.contains('^Unnamed')]
    existing_cols = [c for c in AX_COLUMNS_TO_KEEP if c in df_ax.columns]
    df_ax = df_ax[existing_cols]

    # 3. Filter by RE ID
    re_ids = df_re['Fabric Request Id'].dropna().unique()
    df_ax_filtered = df_ax[df_ax['FTR No'].isin(re_ids)].copy()

    # 4. Create AX PICK
    df_ax_pick = df_ax_filtered.copy()
    df_ax_pick['Roll'] = 1

    # 5. Duplicate Check
    try:
        ws = get_sheet()
        gs_data = pd.DataFrame(ws.get_all_records())
        
        duplicates = []
        if not gs_data.empty and 'FTR No' in gs_data.columns:
            for fid in re_ids:
                if fid in gs_data['FTR No'].values:
                    match = gs_data[gs_data['FTR No'] == fid].iloc[0]
                    duplicates.append(f"⚠️ **Duplicate Found:** ID: {fid} | Date: {match.get('Upload Time','N/A')} | Remark: {match.get('Remark','N/A')}")
        
        if duplicates:
            for d in duplicates: st.warning(d)
        else:
            st.success("✅ සියලුම ID අලුත් ඒවා වේ.")
    except Exception as e:
        st.warning(f"Google Connection Warning: {e}")

    # 6. Advanced Summary Generation
    # RE Summary
    re_summary = df_re.groupby('Fabric Request Id').agg(
        RE_Line_Count=('Fabric Request Id', 'count'),
        RE_Roll_Length_Total=('Roll Length', 'sum')
    ).reset_index()

    # AX Summary
    ax_summary = df_ax_pick.groupby('FTR No').agg(
        AX_Line_Count=('FTR No', 'count'),
        AX_Issue_Qty_Total=('Issue Qty', 'sum')
    ).reset_index()

    # Merge Summary
    df_summary = pd.merge(re_summary, ax_summary, left_on='Fabric Request Id', right_on='FTR No', how='outer')
    df_summary['Status'] = df_summary.apply(
        lambda row: "Matched" if pd.notna(row['FTR No']) and row['RE_Line_Count'] == row['AX_Line_Count'] else "Unmatched", 
        axis=1
    )
    
    # Fill NaN values for cleaner display
    df_summary.fillna("-", inplace=True)
    
    # Rename columns for clarity in output
    df_summary.rename(columns={
        'Fabric Request Id': 'RE Fabric Request ID',
        'FTR No': 'AX FTR No'
    }, inplace=True)

    # Reorder columns
    df_summary = df_summary[['RE Fabric Request ID', 'RE_Roll_Length_Total', 'RE_Line_Count', 'AX FTR No', 'AX_Issue_Qty_Total', 'AX_Line_Count', 'Status']]

    # 7. Dashboard Metrics
    m1, m2, m3, m4 = st.columns(4)
    re_unique_count = len(re_ids)
    re_roll_total = df_re['Roll Length'].sum()
    ax_unique_count = df_ax_pick['FTR No'].nunique()
    ax_qty_total = df_ax_pick['Issue Qty'].sum()

    m1.metric("📌 RE Unique IDs", re_unique_count)
    m2.metric("📏 RE Total Roll Length", round(re_roll_total, 2))
    m3.metric("📌 AX PICK Unique IDs", ax_unique_count)
    m4.metric("📦 AX Total Issue Qty", round(ax_qty_total, 2))

    st.subheader("📊 Summary Preview (Matched/Unmatched)")
    # Highlight Unmatched in Streamlit
    def highlight_unmatched(val):
        color = 'red' if val == 'Unmatched' else 'green'
        return f'color: {color}; font-weight: bold'
    st.dataframe(df_summary.style.map(highlight_unmatched, subset=['Status']), use_container_width=True)

    # 8. Save Data & Excel Generation
    current_time = datetime.now(SL_TIMEZONE).strftime("%Y-%m-%d %I:%M:%S %p")
    
    col_save, col_dl = st.columns(2)
    
    if col_save.button("🚀 Save ALL AX PICK Data to Google Sheet"):
        with st.spinner("Saving to Google Sheets..."):
            try:
                ws = get_sheet()
                # Prepare headers
                gs_headers = ["Upload Time", "Remark"] + list(df_ax_pick.columns)
                
                # If sheet is empty, add headers
                if len(ws.get_all_values()) == 0:
                    ws.append_row(gs_headers)
                else:
                    # Update headers if they mismatch
                    existing_headers = ws.row_values(1)
                    if existing_headers != gs_headers:
                        ws.update('A1', [gs_headers])

                # ✅ ERROR FIX: Replace NaN with empty strings before saving
                df_ax_pick_clean = df_ax_pick.fillna("")
                
                # Prepare data rows
                rows_to_add = []
                for _, row in df_ax_pick_clean.iterrows():
                    row_data = [current_time, remark] + row.tolist()
                    rows_to_add.append(row_data)
                
                ws.append_rows(rows_to_add)
                st.balloons()  
                st.success("AX PICK හි සියලුම දත්ත සාර්ථකව Google Sheet එකට ඇතුළත් කළා!")
            except Exception as e:
                st.error(f"Failed to save: {e}")

    # Prepare Excel for Download
    output = io.BytesIO()
    
    # Add Totals Row to AX PICK for Excel
    df_ax_pick_excel = df_ax_pick.copy()
    total_row = pd.DataFrame([{col: '' for col in df_ax_pick_excel.columns}])
    total_row.loc[0, 'FTR No'] = 'TOTAL'
    total_row.loc[0, 'Issue Qty'] = ax_qty_total
    total_row.loc[0, 'Roll'] = df_ax_pick_excel['Roll'].sum()
    df_ax_pick_excel = pd.concat([df_ax_pick_excel, total_row], ignore_index=True)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ax_pick_excel.to_excel(writer, sheet_name='AX PICK', index=False)
        df_summary.to_excel(writer, sheet_name='Summery', index=False)
        
        workbook = writer.book
        worksheet_ax = writer.sheets['AX PICK']
        worksheet_sum = writer.sheets['Summery']
        
        # Formats
        format_base = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
        format_bold = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'bg_color': '#D9D9D9'})
        format_red = workbook.add_format({'font_name': 'Calibri', 'font_size': 11, 'font_color': '#9C0006', 'bg_color': '#FFC7CE'})
        
        # Apply Base Format
        worksheet_ax.set_column('A:Z', 15, format_base)
        worksheet_sum.set_column('A:Z', 20, format_base)
        
        # Bold Total Row in AX PICK
        last_row_ax = len(df_ax_pick_excel)
        worksheet_ax.set_row(last_row_ax, None, format_bold)

        # Highlight Unmatched in Summary
        worksheet_sum.conditional_format('G2:G1000', {
            'type': 'cell',
            'criteria': '==',
            'value': '"Unmatched"',
            'format': format_red
        })

    col_dl.download_button(
        label="📥 Download Processed Excel",
        data=output.getvalue(),
        file_name=f"FTRN_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- ADMIN EDIT/DELETE DATA VIEW ---
if st.session_state.admin_logged_in:
    st.divider()
    st.subheader("✏️ Edit or Delete Google Sheet Data")
    try:
        ws = get_sheet()
        gs_data = pd.DataFrame(ws.get_all_records())
        if not gs_data.empty:
            st.info("පහත වගුව මත දත්ත වෙනස් කර හෝ මකා දමා 'Save Changes' ඔබන්න.")
            edited_df = st.data_editor(gs_data, num_rows="dynamic", use_container_width=True)
            
            if st.button("💾 Save Changes to Google Sheet"):
                ws.clear()
                # ✅ ERROR FIX: Replace NaN with empty strings in admin editor
                edited_df_clean = edited_df.fillna("")
                ws.update([edited_df_clean.columns.values.tolist()] + edited_df_clean.values.tolist())
                st.toast("දත්ත සාර්ථකව Update විය!", icon="✅")
                time.sleep(1)
                st.rerun()
        else:
            st.write("දැනට Google Sheet එකේ දත්ත නොමැත.")
    except Exception as e:
        st.error(f"Error loading admin data: {e}")

# Footer Branding
st.markdown("<div class='footer'>Developed by Ishanka Madusanka</div>", unsafe_allow_html=True)
