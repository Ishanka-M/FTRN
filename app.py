import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import io

# --- CONFIGURATION ---
SL_TIMEZONE = pytz.timezone('Asia/Colombo')
GOOGLE_SHEET_NAME = "FTRN" 

AX_COLUMNS_TO_KEEP = [
    "Sales Order", "Style No", "FTR No", "Req. Date", "Issued By", 
    "Issued Date", "Item", "Color", "Shade No", "Pattern No", 
    "Serial No", "Batch number", "Unit", "Issue Qty"
]

# --- GOOGLE SHEETS CONNECTION (USING SECRETS) ---
def get_gspread_client():
    # .streamlit/secrets.toml වලින් credentials ලබා ගැනීම
    creds_dict = st.secrets["gcp_service_account"]
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

def rotate_and_clear_sheet(sh):
    """ මධ්‍යම රාත්‍රී 12 ට sheet එක clear කර අලුත් දවසට සූදානම් කිරීම """
    now = datetime.now(SL_TIMEZONE)
    # මධ්‍යම රාත්‍රී 12:00 - 12:15 අතර කාලයේදී sheet එක rotate වේ
    if now.hour == 0 and now.minute <= 15: 
        try:
            main_ws = sh.get_worksheet(0)
            archive_name = f"Archive_{now.strftime('%Y_%m_%d')}"
            
            # පරණ data සහිතව අලුත් sheet එකක් හදන්න
            data = main_ws.get_all_values()
            if len(data) > 1: # Header එකට අමතරව data තිබේ නම් පමණක්
                new_ws = sh.add_worksheet(title=archive_name, rows="1000", cols="20")
                new_ws.update('A1', data)
                
                # ප්‍රධාන sheet එක clear කර headers දාන්න
                main_ws.clear()
                main_ws.append_row(["FTR No", "Upload Time", "Remark", "Issue Qty"])
        except Exception as e:
            st.error(f"Rotation Error: {e}")

# --- STREAMLIT UI ---
st.set_page_config(page_title="FTRN Inventory System", layout="wide")
st.title("📊 FTRN Inventory Control System")

# Sidebar for Inputs
with st.sidebar:
    st.header("Control Panel")
    remark = st.text_input("Remark එක ඇතුළත් කරන්න:")
    uploaded_file = st.file_uploader("Excel File එක තෝරන්න", type=["xlsx"])

if uploaded_file and remark:
    # 1. Load Sheets
    xls = pd.ExcelFile(uploaded_file)
    if 'AX' not in xls.sheet_names or 'RE' not in xls.sheet_names:
        st.error("Error: 'AX' සහ 'RE' sheet දෙකම file එකේ තිබිය යුතුය.")
        st.stop()
        
    df_ax_raw = pd.read_excel(uploaded_file, sheet_name='AX')
    df_re = pd.read_excel(uploaded_file, sheet_name='RE')

    # 2. Clean AX Sheet
    # Headers නැති columns ඉවත් කිරීම (Drop unnamed columns)
    df_ax = df_ax_raw.loc[:, ~df_ax_raw.columns.str.contains('^Unnamed')]
    # අවශ්‍ය columns පමණක් තබා ගැනීම
    existing_cols = [c for c in AX_COLUMNS_TO_KEEP if c in df_ax.columns]
    df_ax = df_ax[existing_cols]

    # 3. Lookup & Filter (Match RE Fabric Request Id with AX FTR No)
    re_ids = df_re['Fabric Request Id'].unique()
    df_ax_filtered = df_ax[df_ax['FTR No'].isin(re_ids)].copy()

    # 4. Create AX PICK Sheet
    df_ax_pick = df_ax_filtered.copy()
    df_ax_pick['Roll'] = 1  # Add Roll column at the end

    # 5. Duplicate Check & Google Sheets
    try:
        client = get_gspread_client()
        sh = client.open(GOOGLE_SHEET_NAME)
        rotate_and_clear_sheet(sh)
        main_ws = sh.get_worksheet(0)
        
        # දැනට Google Sheet එකේ ඇති data කියවීම
        gs_data = pd.DataFrame(main_ws.get_all_records())
        current_time = datetime.now(SL_TIMEZONE).strftime("%Y-%m-%d %I:%M %p")
        
        duplicates = []
        if not gs_data.empty:
            for fid in re_ids:
                if fid in gs_data['FTR No'].values:
                    match = gs_data[gs_data['FTR No'] == fid].iloc[0]
                    duplicates.append(f"⚠️ **Duplicate Found:** ID: {fid} | Date: {match['Upload Time']} | Remark: {match['Remark']}")
        
        if duplicates:
            for d in duplicates: st.warning(d)
        else:
            st.success("✅ සියලුම ID අලුත් ඒවා වේ.")

    except Exception as e:
        st.error(f"Google Connection Error: {e}")

    # 6. Summary Sheet Logic
    re_summary = df_re.groupby('Fabric Request Id').agg({
        'Roll Length': 'sum', 
        'Fabric Request Id': 'count'
    }).rename(columns={'Fabric Request Id': 'Line Count'}).reset_index()
    
    summary_data = []
    for _, row in re_summary.iterrows():
        fid = row['Fabric Request Id']
        ax_match = df_ax_pick[df_ax_pick['FTR No'] == fid]
        
        status = "Matched" if not ax_match.empty and len(ax_match) == row['Line Count'] else "Unmatched"
        summary_data.append({
            "Fabric Request Id": fid,
            "Total Roll Length": row['Roll Length'],
            "Line Count": row['Line Count'],
            "Status": status
        })
    df_summary = pd.DataFrame(summary_data)

    # 7. Dashboard
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("RE Unique IDs", len(re_ids))
    m2.metric("RE Total Roll Length", round(df_re['Roll Length'].sum(), 2))
    m3.metric("AX PICK Unique FTR", df_ax_pick['FTR No'].nunique())
    m4.metric("AX Total Issue Qty", round(df_ax_pick['Issue Qty'].sum(), 2))

    # 8. Action Buttons
    col_save, col_dl = st.columns(2)
    
    if col_save.button("🚀 Save Data to Google Sheet"):
        # FTR No එක පදනම් කරගෙන Unique rows පමණක් save කිරීම
        rows_to_add = []
        unique_ax_pick = df_ax_pick.drop_duplicates(subset=['FTR No'])
        for _, row in unique_ax_pick.iterrows():
            rows_to_add.append([row['FTR No'], current_time, remark, row['Issue Qty']])
        
        main_ws.append_rows(rows_to_add)
        st.success("දත්ත FTRN sheet එකට සාර්ථකව ඇතුළත් කළා!")

    # Excel Download (Calibri 11)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ax_pick.to_excel(writer, sheet_name='AX PICK', index=False)
        df_summary.to_excel(writer, sheet_name='Summery', index=False)
        
        workbook = writer.book
        fmt = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
        for sn in ['AX PICK', 'Summery']:
            ws = writer.sheets[sn]
            ws.set_column('A:Z', None, fmt)

    col_dl.download_button(
        label="📥 Download Processed Excel",
        data=output.getvalue(),
        file_name=f"FTRN_Output_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("කරුණාකර Excel file එක upload කර Remark එකක් ඇතුළත් කරන්න.")
