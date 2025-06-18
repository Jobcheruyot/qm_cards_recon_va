# Regenerate the Streamlit app with read_excel engine fix and defensive error handling
streamlit_fixed_script_path = "/mnt/data/card_reconciliation_app_fixed.py"

streamlit_script_with_fix = '''\
import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Card Reconciliation App", layout="wide")

st.title("🧾 Card Reconciliation Processor")

# --- File uploads ---
kcb_file = st.file_uploader("Upload KCB Excel File", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel File", type=["xlsx"])
aspire_file = st.file_uploader("Upload Aspire CSV File", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel File", type=["xlsx"])

def safe_read_excel(file, name):
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Failed to read {name} Excel file: {e}")
        st.stop()

if all([kcb_file, equity_file, aspire_file, key_file]):
    kcb = safe_read_excel(kcb_file, "KCB")
    equity = safe_read_excel(equity_file, "Equity")
    aspire = pd.read_csv(aspire_file)
    card_key = safe_read_excel(key_file, "Card Key")

    # === KCB Cleaning ===
    kcb.columns = kcb.columns.str.strip()
    equity.columns = equity.columns.str.strip()
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

    equity = equity.rename(columns={'Outlet_Name': 'store'})
    equity['Source'] = 'Equity'

    columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
               'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
    kcb_final = kcb_renamed[columns]
    equity_final = equity[columns]
    merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)

    merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
    merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']

    card_key.columns = card_key.columns.str.strip()
    card_key['Col_1'] = card_key['Col_1'].str.strip()
    card_key['Col_2'] = card_key['Col_2'].str.strip()
    merged_cards['store'] = merged_cards['store'].str.strip()
    lookup_dict = dict(zip(card_key['Col_1'], card_key['Col_2']))
    merged_cards['branch'] = merged_cards['store'].map(lookup_dict)
    merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
    merged_cards['card_check'] = merged_cards['Card_Number'].apply(lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else '')
    merged_cards = merged_cards.drop_duplicates()
    merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]

    # === Aspire cleaning ===
    aspire['CARD_NUMBER'] = aspire['CARD_NUMBER'].astype(str).str.strip()
    aspire['card_check'] = aspire['CARD_NUMBER'].apply(lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else '')
    aspire = aspire[[
        'STORE_CODE', 'STORE_NAME', 'ZED_DATE', 'TILL', 'SESSION',
        'RCT', 'CUSTOMER_NAME', 'CARD_TYPE', 'CARD_NUMBER', 'card_check',
        'AMOUNT', 'REF_NO', 'RCT_TRN_DATE'
    ]]
    aspire = aspire.rename(columns={'REF_NO': 'R_R_N'})
    aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
    merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()
    rrntable = pd.merge(aspire, merged_cards, on='R_R_N', how='inner', suffixes=('_aspire', '_merged'))

    # --- card_summary ---
    card_summary = (
        aspire['STORE_NAME']
        .dropna()
        .drop_duplicates()
        .sort_values()
        .reset_index(drop=True)
        .to_frame(name='STORE_NAME')
    )
    card_summary.index = card_summary.index + 1
    card_summary.reset_index(inplace=True)
    card_summary.rename(columns={'index': 'No'}, inplace=True)

    aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
    aspire_sums = aspire.groupby('STORE_NAME')['AMOUNT'].sum().reset_index().rename(columns={'AMOUNT': 'Aspire_Zed'})
    card_summary = card_summary.merge(aspire_sums, on='STORE_NAME', how='left').fillna(0)

    merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')
    kcb_grouped = merged_cards[merged_cards['Source'] == 'KCB'].groupby('branch')['Purchase'].sum().reset_index().rename(columns={'branch': 'STORE_NAME', 'Purchase': 'kcb_paid'})
    equity_grouped = merged_cards[merged_cards['Source'] == 'Equity'].groupby('branch')['Purchase'].sum().reset_index().rename(columns={'branch': 'STORE_NAME', 'Purchase': 'equity_paid'})
    card_summary = card_summary.merge(kcb_grouped, on='STORE_NAME', how='left').merge(equity_grouped, on='STORE_NAME', how='left').fillna(0)

    card_summary['Gross_Banking'] = card_summary['kcb_paid'] + card_summary['equity_paid']
    total_row = pd.DataFrame([{
        'No': '',
        'STORE_NAME': 'TOTAL',
        'Aspire_Zed': card_summary['Aspire_Zed'].sum(),
        'kcb_paid': card_summary['kcb_paid'].sum(),
        'equity_paid': card_summary['equity_paid'].sum(),
        'Gross_Banking': card_summary['Gross_Banking'].sum()
    }])
    card_summary = pd.concat([card_summary, total_row], ignore_index=True)

    # --- Prepare Output File ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        card_summary.to_excel(writer, sheet_name='card_summary', index=False)
        rrntable.to_excel(writer, sheet_name='Aspire_recs', index=False)
        merged_cards[merged_cards['Source'] == 'KCB'].to_excel(writer, sheet_name='KCB_recs', index=False)
        merged_cards[merged_cards['Source'] == 'Equity'].to_excel(writer, sheet_name='Equity_recs', index=False)
        merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
        aspire.to_excel(writer, sheet_name='newaspire', index=False)
    output.seek(0)

    # --- Download button ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.success("✅ Reconciliation complete. Click below to download.")
    st.download_button(
        label="📥 Download Reconciliation Workbook",
        data=output,
        file_name=f"Card_Reconciliation_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("👆 Please upload all four files to proceed.")
'''

# Save the fixed script
with open(streamlit_fixed_script_path, "w", encoding="utf-8") as f:
    f.write(streamlit_script_with_fix)

streamlit_fixed_script_path
