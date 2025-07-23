import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Card Reconciliation Report", layout="wide")
st.title("🗞️ Card Reconciliation Advanced Processor")

# --- File Uploads ---
kcb_file = st.file_uploader("Upload KCB Excel File", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel File", type=["xlsx"])
coop_file = st.file_uploader("Upload Coop Excel File", type=["xls"])
aspire_file = st.file_uploader("Upload Aspire CSV File", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel File", type=["xlsx"])

# --- Helper functions ---
def safe_read_excel(file, name):
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to read {name} Excel file: {e}")
        st.stop()

def safe_read_csv(file, name):
    try:
        try:
            return pd.read_csv(file)
        except pd.errors.ParserError:
            file.seek(0)
            return pd.read_csv(file, on_bad_lines='skip')
    except Exception as e:
        st.error(f"❌ Failed to read {name} CSV file: {e}")
        st.stop()

# --- Main Processing ---
if all([kcb_file, equity_file, coop_file, aspire_file, key_file]):
    # Read all files
    kcb = safe_read_excel(kcb_file, "KCB")
    equity = safe_read_excel(equity_file, "Equity")
    coop_raw = pd.read_excel(coop_file, skiprows=5)
    aspire = safe_read_csv(aspire_file, "Aspire")
    key = safe_read_excel(key_file, "Card Key")

    # Standardize KCB
    kcb.columns = kcb.columns.str.strip()
    kcb_renamed = kcb.rename(columns={
        'Card No': 'Card_Number',
        'Trans Date': 'TRANS_DATE',
        'RRN': 'R_R_N',
        'Amount': 'Purchase',
        'Comm': 'Commission',
        'NetPaid': 'Settlement_Amount',
        'Merchant': 'store'
    })
    kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -x if x < 0 else 0)
    kcb_renamed['Source'] = 'KCB'

    # Standardize Equity
    equity.columns = equity.columns.str.strip()
    equity = equity.rename(columns={'Outlet_Name': 'store'})
    equity['Source'] = 'Equity'

    # Clean Coop
    coop_cleaned = coop_raw.dropna(subset=["TRANSACTION DATE"]).copy()
    coop_cleaned['Source'] = 'coop'
    coop_map = {

    "TERMINAL": "TERMINAL",
    "LOCATION": "store",
    "CARD": "Card_Number",
    "TRANSACTION DATE": "TRANS_DATE",
    "RRN CODE": "R_R_N",
    "TRANSACTION AMOUNT": "Purchase",
    "BANK COMM": "Commission",
    "NET PAID": "Settlement_Amount",
    "CASH BACK": "Cash_Back",
    "Source": "Source"
    }
    coop_selected = coop_cleaned[list(coop_map.keys())].rename(columns=coop_map)

    # --- Column Alignment ---
    columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
               'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
    kcb_final = kcb_renamed[columns]
    equity_final = equity[columns]

    merged_cards = pd.concat([kcb_final, equity_final, coop_selected], ignore_index=True)
    merged_cards = merged_cards[merged_cards['Card_Number'].notna() & (merged_cards['Card_Number'].astype(str).str.strip() != '')]

    # Card Key Mapping
    key.columns = key.columns.str.strip()
    key['Col_1'] = key['Col_1'].astype(str).str.strip()
    key['Col_2'] = key['Col_2'].astype(str).str.strip()
    lookup_dict = dict(zip(key['Col_1'], key['Col_2']))
    merged_cards['store'] = merged_cards['store'].astype(str).str.strip()
    merged_cards.loc[merged_cards['Source'] == 'coop', 'branch'] = merged_cards.loc[merged_cards['Source'] == 'coop', 'store'].map(lookup_dict)
    merged_cards['branch'] = merged_cards['branch'].fillna(merged_cards['store'].map(lookup_dict))
    merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
    merged_cards['card_check'] = merged_cards['Card_Number'].apply(lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else '')
    merged_cards = merged_cards.drop_duplicates()

    # Final preview
    st.subheader("📊 Preview Merged Card Data")
    st.dataframe(merged_cards.head())

    # Option to export merged_cards
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
    output.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.download_button(
        label="📅 Download Merged_Cards.xlsx",
        data=output,
        file_name=f"Merged_Cards_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("👇 Please upload all five files to continue.")
