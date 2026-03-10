import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import pytz
import io

# --- CONFIGURATION ---
SL_TIMEZONE = pytz.timezone('Asia/Colombo')
GOOGLE_SHEET_NAME = "FTRN"  # ඔබ සඳහන් කළ Sheet එකේ නම
JSON_KEYFILE = "service_account.json" 

# තබාගත යුතු Columns ලැයිස්තුව
AX_COLUMNS_TO_KEEP = [
    "Sales Order", "Style No", "FTR No", "Req. Date", "Issued By", 
    "Issued Date", "Item", "Color", "Shade No", "Pattern No", 
    "Serial No", "Batch number", "Unit", "Issue Qty"
]

# --- GOOGLE SHEETS CONNECTION ---
def get_gspread_client():
    # Service account එක හරහා සම්බන්ධ වීම
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEYFILE, scope)
    return gspread.authorize(creds)

def rotate_and_clear_sheet(sh):
    """ මධ්‍යම රාත්‍රී 12 ට sheet එක clear කර අලුත් දවසට සූදානම් කිරීම """
    now = datetime.now(SL_TIMEZONE)
    if now.hour == 0 and now.minute <= 10: # 12:00 - 12:10 අතර පමණක්
        try:
            main_ws = sh.get_worksheet(0)
            archive_name = f"Archive_{now.strftime('%Y_%m_%d')}"
            
            # පරණ data සහිතව අලුත් sheet එකක් හදන්න
            data = main_ws.get_all_values()
            new_ws = sh.add_worksheet(title=archive_name, rows="1000", cols="20")
            new_ws.update('A1', data)
            
            # ප්‍රධාන sheet එක clear කර headers දාන්න
            main_ws.clear()
            main_ws.append_row(["FTR No", "Upload Time", "Remark", "Issue Qty"])
        except Exception as e:
            st.error(f"Automation Error: {e}")

# --- UI SETUP ---
st.set_page_config(page_title="Inventory Control - FTRN", layout="wide")
st.title("📦 FTRN Inventory Processing System")

with st.sidebar:
    st.header("Upload Settings")
    remark = st.text_input("Remark එක ඇතුළත් කරන්න:")
    uploaded_file = st.file_uploader("Excel File එක තෝරන්න", type=["xlsx"])

if uploaded_file and remark:
    # 1. දත්ත කියවීම (Reading Data)
    df_ax_raw = pd.read_excel(uploaded_file, sheet_name='AX')
    df_re = pd.read_excel(uploaded_file, sheet_name='RE')

    # 2. AX Sheet පිරිසිදු කිරීම
    # Headers නැති columns ඉවත් කිරීම
    df_ax = df_ax_raw.loc[:, ~df_ax_raw.columns.str.contains('^Unnamed')]
    # අවශ්‍ය columns පමණක් තබා ගැනීම
    df_ax = df_ax[AX_COLUMNS_TO_KEEP]

    # 3. RE sheet එකේ ID සමඟ AX filter කිරීම
    re_ids = df_re['Fabric Request Id'].unique()
    df_ax_filtered = df_ax[df_ax['FTR No'].isin(re_ids)].copy()

    # 4. AX PICK Sheet එක සැකසීම
    df_ax_pick = df_ax_filtered.copy()
    df_ax_pick['Roll'] = 1  # අන්තිමට Roll column එක එකතු කිරීම

    # 5. Google Sheet එක පරීක්ෂා කිරීම (Duplicate Check)
    try:
        client = get_gspread_client()
        sh = client.open(GOOGLE_SHEET_NAME)
        rotate_and_clear_sheet(sh)
        main_ws = sh.get_worksheet(0)
        
        # දැනට තිබෙන දත්ත ලබා ගැනීම
        existing_data = pd.DataFrame(main_ws.get_all_records())
        current_time = datetime.now(SL_TIMEZONE).strftime("%Y-%m-%d %I:%M %p")
        
        if not existing_data.empty:
            duplicate_found = False
            for fid in re_ids:
                if fid in existing_data['FTR No'].values:
                    match_row = existing_data[existing_data['FTR No'] == fid].iloc[0]
                    st.warning(f"⚠️ **Duplicate ID:** {fid} | කලින් දාපු වෙලාව: {match_row['Upload Time']} | Remark: {match_row['Remark']}")
                    duplicate_found = True
            
            if not duplicate_found:
                st.success("✅ මෙම File එකේ ඇති සියලුම ID අලුත් ඒවා වේ.")
    except Exception as e:
        st.error(f"Google Sheet Error: {e}")

    # 6. Summary Sheet සැකසීම
    re_summary = df_re.groupby('Fabric Request Id').agg({'Roll Length': 'sum', 'Fabric Request Id': 'count'}).rename(columns={'Fabric Request Id': 'Line Count'}).reset_index()
    
    summary_rows = []
    for _, row in re_summary.iterrows():
        fid = row['Fabric Request Id']
        ax_match = df_ax_pick[df_ax_pick['FTR No'] == fid]
        
        # Matching Logic
        status = "Matched" if not ax_match.empty and len(ax_match) == row['Line Count'] else "Unmatched"
        summary_rows.append({
            "Fabric Request Id": fid,
            "Total Roll Length": row['Roll Length'],
            "Line Count": row['Line Count'],
            "Status": status
        })
    df_summary = pd.DataFrame(summary_rows)

    # 7. Dashboard Metrics
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("RE Total IDs", len(re_ids))
    m2.metric("RE Total Roll Length", f"{df_re['Roll Length'].sum():.2f}")
    m3.metric("AX Unique FTR", df_ax_pick['FTR No'].nunique())
    m4.metric("AX Total Issue Qty", f"{df_ax_pick['Issue Qty'].sum():.2f}")

    # 8. Save & Download
    col_save, col_dl = st.columns(2)
    
    if col_save.button("Save to Google Sheet (FTRN)"):
        new_entries = []
        for _, row in df_ax_pick.drop_duplicates(subset=['FTR No']).iterrows():
            new_entries.append([row['FTR No'], current_time, remark, row['Issue Qty']])
        
        main_ws.append_rows(new_entries)
        st.balloons()
        st.info("දත්ත සාර්ථකව FTRN sheet එකට ඇතුළත් කරන ලදී!")

    # Excel formatting (Calibri, Size 11)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_ax_pick.to_excel(writer, sheet_name='AX PICK', index=False)
        df_summary.to_excel(writer, sheet_name='Summery', index=False)
        
        workbook = writer.book
        cell_format = workbook.add_format({'font_name': 'Calibri', 'font_size': 11})
        
        for sheet_name in ['AX PICK', 'Summery']:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:Z', None, cell_format)

    col_dl.download_button(
        label="Download Processed Excel",
        data=output.getvalue(),
        file_name=f"FTRN_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("කරුණාකර Excel file එකක් Upload කර Remark එකක් ඇතුළත් කරන්න.")
