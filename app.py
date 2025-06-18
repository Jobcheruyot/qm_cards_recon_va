import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Card Reconciliation Report", layout="wide")
st.title("🧾 Card Reconciliation Advanced Processor")

# --- File Uploads ---
kcb_file = st.file_uploader("Upload KCB Excel File", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel File", type=["xlsx"])
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
        # Try to read with default settings
        try:
            return pd.read_csv(file)
        except pd.errors.ParserError as e:
            # Try again skipping bad lines if error occurs
            file.seek(0)  # Reset file pointer
            df = pd.read_csv(file, on_bad_lines='skip')
            st.warning(f"⚠️ Some rows in {name} CSV were skipped due to formatting issues.")
            return df
    except Exception as e:
        st.error(f"❌ Failed to read {name} CSV file: {e}")
        st.stop()

# --- Main Processing ---
if all([kcb_file, equity_file, aspire_file, key_file]):
    # --- Read files ---
    kcb = safe_read_excel(kcb_file, "KCB")
    equity = safe_read_excel(equity_file, "Equity")
    aspire = safe_read_csv(aspire_file, "Aspire")
    key = safe_read_excel(key_file, "Card Key")

    # --- Bank data alignment ---
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
    merged_cards = merged_cards[merged_cards['Card_Number'].notna() & (merged_cards['Card_Number'].astype(str).str.strip() != '')]

    # --- Card Key Map ---
    card_key = key.copy()
    card_key.columns = card_key.columns.str.strip()
    card_key['Col_1'] = card_key['Col_1'].astype(str).str.strip()
    card_key['Col_2'] = card_key['Col_2'].astype(str).str.strip()
    merged_cards['store'] = merged_cards['store'].astype(str).str.strip()
    lookup_dict = dict(zip(card_key['Col_1'], card_key['Col_2']))
    merged_cards['branch'] = merged_cards['store'].map(lookup_dict)
    merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
    merged_cards['card_check'] = merged_cards['Card_Number'].apply(lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else '')
    merged_cards = merged_cards.drop_duplicates()
    merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]

    # --- Aspire cleaning ---
    aspire['CARD_NUMBER'] = aspire['CARD_NUMBER'].astype(str).str.strip()
    aspire['card_check'] = aspire['CARD_NUMBER'].apply(lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else '')
    aspire = aspire[[
        'STORE_CODE', 'STORE_NAME', 'ZED_DATE', 'TILL', 'SESSION',
        'RCT', 'CUSTOMER_NAME', 'CARD_TYPE', 'CARD_NUMBER', 'card_check',
        'AMOUNT', 'REF_NO', 'RCT_TRN_DATE'
    ]].copy()
    aspire = aspire.rename(columns={'REF_NO': 'R_R_N'})
    aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
    merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()

    # --- Merge on R_R_N for rrntable ---
    rrntable = pd.merge(aspire, merged_cards, on='R_R_N', how='inner', suffixes=('_aspire', '_merged'))

    # --- Card Summary Table ---
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
    card_summary['Aspire_Zed'] = pd.to_numeric(card_summary['Aspire_Zed'], errors='coerce')
    card_summary['kcb_paid'] = pd.to_numeric(card_summary['kcb_paid'], errors='coerce')
    card_summary['equity_paid'] = pd.to_numeric(card_summary['equity_paid'], errors='coerce')
    card_summary['Gross_Banking'] = pd.to_numeric(card_summary['Gross_Banking'], errors='coerce')

    # --- Recs logic for unmatched/variance analysis ---
    # Step 1: RRN check (aspire vs merged_cards)
    ref_to_purchase = dict(zip(merged_cards['R_R_N'], merged_cards['Purchase']))
    aspire['rrn_check'] = aspire['R_R_N'].map(ref_to_purchase).fillna(0)
    aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']

    # Mark aspire unmatched by RRN (<=0)
    newaspire = aspire[aspire['rrn_check'] <= 0].copy()

    # Generate newbankmerged: merged_cards not matched in aspire by RRN
    matched_ref_nos = set(aspire['R_R_N'])
    merged_cards['Cheked_rows'] = merged_cards['R_R_N'].astype(str).apply(lambda x: 'Yes' if x in matched_ref_nos else 'No')
    newbankmerged = merged_cards[merged_cards['Cheked_rows'].str.strip().str.upper() == 'NO'].copy()

    # Amount-check logic for newaspire and newbankmerged
    newaspire['Check_Two'] = newaspire['STORE_NAME'].astype(str).str.replace(r'\s+', '', regex=True).str.upper() + newaspire['AMOUNT'].astype(float).astype(int).astype(str)
    newbankmerged['branch'] = newbankmerged['branch'].astype(str)
    newbankmerged['Purchase'] = pd.to_numeric(newbankmerged['Purchase'], errors='coerce')
    newbankmerged['Check_Two'] = newbankmerged['branch'].str.replace(r'\s+', '', regex=True).str.upper() + newbankmerged['Purchase'].astype(float).astype(int).astype(str)

    # Unique match logic (consume once)
    available_matches = newbankmerged['Check_Two'].tolist()
    def check_and_consume(val):
        if val in available_matches:
            available_matches.remove(val)
            return 'Okay'
        else:
            return 'False'
    newaspire['Amount_check'] = newaspire['Check_Two'].apply(check_and_consume)

    # Same for newbankmerged vs newaspire
    aspire_available_matches = newaspire['Check_Two'].tolist()
    def check_and_consume_from_aspire(val):
        if val in aspire_available_matches:
            aspire_available_matches.remove(val)
            return 'Okay'
        else:
            return 'False'
    newbankmerged['Amount_check'] = newbankmerged['Check_Two'].apply(check_and_consume_from_aspire)

    # newmerged_cards: truly unmatched bank records after both checks
    newmerged_cards = newbankmerged[
        (newbankmerged['Cheked_rows'].str.strip().str.lower() == 'no') &
        (newbankmerged['Amount_check'].astype(str).str.strip().str.lower() == 'false')
    ].copy()

    # --- Final recs for reporting ---
    aspire_recs_data = newaspire[newaspire['Amount_check'] == 'False'].copy()
    equity_recs_data = newmerged_cards[
        (newmerged_cards['Source'].str.upper() == 'EQUITY') &
        (newmerged_cards['Amount_check'] == 'False')
    ].copy()
    kcb_recs_data = newmerged_cards[
        (newmerged_cards['Source'].str.upper() == 'KCB') &
        (newmerged_cards['Amount_check'] == 'False')
    ].copy()

    # --- Add summary makeup row, variance, etc ---
    # Remove old TOTAL row if any
    card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']

    # Add recs columns (sum by branch/STORE_NAME)
    kcb_recs_grouped = kcb_recs_data.groupby('branch')['Purchase'].sum().reset_index()
    kcb_recs_grouped.columns = ['STORE_NAME', 'kcb_recs']
    card_summary = card_summary.merge(kcb_recs_grouped, on='STORE_NAME', how='left')

    equity_recs_grouped = equity_recs_data.groupby('branch')['Purchase'].sum().reset_index()
    equity_recs_grouped.columns = ['STORE_NAME', 'Equity_recs']
    card_summary = card_summary.merge(equity_recs_grouped, on='STORE_NAME', how='left')

    aspire_recs_grouped = aspire_recs_data.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
    aspire_recs_grouped.columns = ['STORE_NAME', 'Asp_Recs']
    card_summary = card_summary.merge(aspire_recs_grouped, on='STORE_NAME', how='left')

    for col in ['kcb_recs', 'Equity_recs', 'Asp_Recs']:
        if col not in card_summary.columns:
            card_summary[col] = 0
        card_summary[col] = pd.to_numeric(card_summary[col], errors='coerce').fillna(0)

    # --- Variance and Net_variance ---
    card_summary['Variance'] = card_summary['Gross_Banking'] - card_summary['Aspire_Zed']
    card_summary['Net_variance'] = card_summary['Variance'] - card_summary['kcb_recs'] - card_summary['Equity_recs'] + card_summary['Asp_Recs']

    # --- Add TOTAL row ---
    numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                    'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs', 'Net_variance']
    totals = card_summary[numeric_cols].sum()
    total_row = pd.DataFrame([{
        'No': '',
        'STORE_NAME': 'TOTAL',
        **{col: totals[col] for col in numeric_cols}
    }])
    card_summary = pd.concat([card_summary, total_row], ignore_index=True)

    # --- REARRANGE COLUMNS AS REQUESTED ---
    card_summary = card_summary[
        ['No', 'STORE_NAME', 'Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance',
         'kcb_recs', 'Equity_recs', 'Asp_Recs', 'Net_variance']
    ]

    # --- Export workbook ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        card_summary.to_excel(writer, sheet_name='card_summary', index=False)
        aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)
        equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)
        kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)
        merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
        aspire.to_excel(writer, sheet_name='aspire', index=False)
        rrntable.to_excel(writer, sheet_name='Aspire_recs', index=False)
    output.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.success("✅ Reconciliation complete. Click below to download the full advanced workbook.")
    st.download_button(
        label="📥 Download Reconciliation_Report.xlsx",
        data=output,
        file_name=f"Reconciliation_Report_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.warning("👆 Please upload all four files to proceed.")
