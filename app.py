import pandas as pd
import streamlit as st

st.set_page_config(layout="wide", page_title="Card Reconciliation Report", initial_sidebar_state="expanded")

st.title("Card Reconciliation Report (Streamlit Version)")

with st.sidebar:
    st.header("Upload Data Files")
    kcb_file = st.file_uploader("KCB Excel", type=["xlsx"])
    equity_file = st.file_uploader("Equity Excel", type=["xlsx"])
    aspire_file = st.file_uploader("Aspire CSV", type=["csv"])
    cardkey_file = st.file_uploader("Card Key Excel", type=["xlsx"])
    st.markdown("---")
    st.info("Upload all required files to enable reconciliation.")

if not (kcb_file and equity_file and aspire_file and cardkey_file):
    st.warning("Please upload all required files (KCB, Equity, Aspire, Card Key) to proceed.")
    st.stop()

# -------------------- LOAD FILES --------------------
kcb = pd.read_excel(kcb_file)
equity = pd.read_excel(equity_file)
aspire = pd.read_csv(aspire_file)
key = pd.read_excel(cardkey_file)

# -------------------- CLEAN COLUMNS --------------------
kcb.columns = kcb.columns.str.strip()
equity.columns = equity.columns.str.strip()
key.columns = key.columns.str.strip()

# -------------------- KCB PROCESSING --------------------
kcb_renamed = kcb.rename(columns={
    'Card No': 'Card_Number',
    'Trans Date': 'TRANS_DATE',
    'RRN': 'R_R_N',
    'Amount': 'Purchase',
    'Comm': 'Commission',
    'NetPaid': 'Settlement_Amount',
    'Merchant': 'store'
})
kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -1 * x if x < 0 else 0)
kcb_renamed['Source'] = 'KCB'

# -------------------- EQUITY PROCESSING --------------------
equity = equity.rename(columns={'Outlet_Name': 'store'})
equity['Source'] = 'Equity'

columns = [
    'TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
    'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source'
]
kcb_final = kcb_renamed[columns]
equity_final = equity[columns]

merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)
merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']

# -------------------- MAP BRANCHES VIA CARD KEY --------------------
key['Col_1'] = key['Col_1'].astype(str).str.strip()
key['Col_2'] = key['Col_2'].astype(str).str.strip()
merged_cards['store'] = merged_cards['store'].astype(str).str.strip()
lookup_dict = dict(zip(key['Col_1'], key['Col_2']))
merged_cards['branch'] = merged_cards['store'].map(lookup_dict)

cols = list(merged_cards.columns)
source_index = cols.index('Source')
if 'branch' in cols:
    cols.insert(source_index + 1, cols.pop(cols.index('branch')))
merged_cards = merged_cards[cols]

# -------------------- ADD CARD CHECK COLUMN --------------------
merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
merged_cards['card_check'] = merged_cards['Card_Number'].apply(
    lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
)
if 'branch' in merged_cards.columns and 'card_check' in merged_cards.columns:
    cols = merged_cards.columns.tolist()
    cols.remove('card_check')
    branch_index = cols.index('branch')
    cols.insert(branch_index + 1, 'card_check')
    merged_cards = merged_cards[cols]

# -------------------- ASPIRE PROCESSING --------------------
aspire['CARD_NUMBER'] = aspire['CARD_NUMBER'].astype(str).str.strip()
aspire['card_check'] = aspire['CARD_NUMBER'].apply(
    lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
)
cols = aspire.columns.tolist()
if 'CARD_NUMBER' in cols and 'card_check' in cols:
    cols.remove('card_check')
    insert_index = cols.index('CARD_NUMBER') + 1
    cols.insert(insert_index, 'card_check')
    aspire = aspire[cols]

aspire_cols = [
    'STORE_CODE','STORE_NAME','ZED_DATE','TILL','SESSION','RCT','CUSTOMER_NAME',
    'CARD_TYPE','CARD_NUMBER','card_check','AMOUNT','REF_NO','RCT_TRN_DATE'
]
aspire = aspire[[c for c in aspire_cols if c in aspire.columns]]
aspire = aspire.rename(columns={'REF_NO': 'R_R_N'})

aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()

# -------------------- NEWMERGED_CARDS FOR RECS LOGIC --------------------
# For the recs logic, you need Amount_check for unmatched logic
# For demonstration, mark all as unmatched (replace with your actual matching logic as needed)
newmerged_cards = merged_cards.copy()
newmerged_cards['Amount_check'] = 'False'

# -------------------- NEWASPIRE FOR RECS LOGIC --------------------
newaspire = aspire.copy()
newaspire['Amount_check'] = 'False'  # Replace with real logic if you have it

# -------------------- CARD SUMMARY --------------------
aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')

card_summary = aspire['STORE_NAME'].dropna().drop_duplicates().sort_values().reset_index(drop=True).to_frame(name='STORE_NAME')
card_summary.index = card_summary.index + 1
card_summary.reset_index(inplace=True)
card_summary.rename(columns={'index': 'No'}, inplace=True)

aspire_sums = aspire.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
aspire_sums = aspire_sums.rename(columns={'AMOUNT': 'Aspire_Zed'})
card_summary = card_summary.merge(aspire_sums, on='STORE_NAME', how='left')
card_summary['Aspire_Zed'] = card_summary['Aspire_Zed'].fillna(0)

kcb_grouped = merged_cards[merged_cards['Source'] == 'KCB'].groupby('branch')['Purchase'].sum().reset_index().rename(columns={'branch': 'STORE_NAME', 'Purchase': 'kcb_paid'})
equity_grouped = merged_cards[merged_cards['Source'] == 'Equity'].groupby('branch')['Purchase'].sum().reset_index().rename(columns={'branch': 'STORE_NAME', 'Purchase': 'equity_paid'})
card_summary = card_summary.merge(kcb_grouped, on='STORE_NAME', how='left')
card_summary['kcb_paid'] = card_summary['kcb_paid'].fillna(0)
card_summary = card_summary.merge(equity_grouped, on='STORE_NAME', how='left')
card_summary['equity_paid'] = card_summary['equity_paid'].fillna(0)

card_summary['Gross_Banking'] = card_summary['kcb_paid'] + card_summary['equity_paid']
for col in ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking']:
    card_summary[col] = card_summary[col].astype(float)
card_summary['Variance'] = card_summary['Gross_Banking'] - card_summary['Aspire_Zed']

# ========== RECS LOGIC ==========

def update_recs_column(card_summary, group_df, on_col, recs_col):
    if recs_col in card_summary.columns:
        card_summary = card_summary.drop(columns=[recs_col])
    card_summary = card_summary[card_summary[on_col] != 'TOTAL'].copy()
    card_summary = card_summary.merge(group_df, on=on_col, how='left')
    card_summary[recs_col] = card_summary[recs_col].fillna(0)
    return card_summary

def append_total_row(card_summary, on_col):
    card_summary = card_summary[card_summary[on_col] != 'TOTAL'].copy()
    numeric_cols = card_summary.select_dtypes('number').columns.tolist()
    totals = card_summary[numeric_cols].sum()
    total_row = {on_col: 'TOTAL'}
    for col in card_summary.columns:
        total_row[col] = totals[col] if col in totals else ''
    card_summary = pd.concat([card_summary, pd.DataFrame([total_row])], ignore_index=True)
    return card_summary

# --- KCB recs ---
kcb_recs_data = newmerged_cards[
    (newmerged_cards['Source'].str.upper() == 'KCB') &
    (newmerged_cards['Amount_check'].astype(str).str.strip().str.lower() == 'false')
].copy()
kcb_recs_data['Purchase'] = pd.to_numeric(kcb_recs_data['Purchase'], errors='coerce')
kcb_recs_data = kcb_recs_data.dropna(subset=['Purchase'])
kcb_recs_grouped = kcb_recs_data.groupby('branch')['Purchase'].sum().reset_index()
kcb_recs_grouped.columns = ['STORE_NAME', 'kcb_recs']
card_summary = update_recs_column(card_summary, kcb_recs_grouped, 'STORE_NAME', 'kcb_recs')

# --- Equity recs ---
equity_recs_data = newmerged_cards[
    (newmerged_cards['Source'].str.upper() == 'EQUITY') &
    (newmerged_cards['Amount_check'].astype(str).str.strip().str.lower() == 'false')
].copy()
equity_recs_data['Purchase'] = pd.to_numeric(equity_recs_data['Purchase'], errors='coerce')
equity_recs_data = equity_recs_data.dropna(subset=['Purchase'])
equity_recs_grouped = equity_recs_data.groupby('branch')['Purchase'].sum().reset_index()
equity_recs_grouped.columns = ['STORE_NAME', 'Equity_recs']
card_summary = update_recs_column(card_summary, equity_recs_grouped, 'STORE_NAME', 'Equity_recs')

# --- Aspire recs ---
aspire_recs_data = newaspire[
    newaspire['Amount_check'].astype(str).str.strip().str.lower() == 'false'
].copy()
aspire_recs_data['AMOUNT'] = pd.to_numeric(aspire_recs_data['AMOUNT'], errors='coerce')
aspire_recs_data = aspire_recs_data.dropna(subset=['AMOUNT'])
aspire_recs_grouped = aspire_recs_data.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
aspire_recs_grouped.columns = ['STORE_NAME', 'Asp_Recs']
card_summary = update_recs_column(card_summary, aspire_recs_grouped, 'STORE_NAME', 'Asp_Recs')

# --- Append new TOTAL row ---
card_summary = append_total_row(card_summary, 'STORE_NAME')

# ----------- FINAL MULTI-SHEET RECONCILIATION EXPORT -----------
@st.cache_data
def to_excel_final_report(**dfs):
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=sheet)
    output.seek(0)
    return output.getvalue()

st.download_button(
    "Download FULL Reconciliation_Report.xlsx (all sheets)",
    data=to_excel_final_report(
        card_summary=card_summary,
        KCB_recs=kcb_recs_data,
        Equity_recs=equity_recs_data,
        Aspire_recs=aspire_recs_data,
        merged_cards=merged_cards,
        newaspire=newaspire
    ),
    file_name="Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("Card Summary (Final Preview)"):
    st.dataframe(card_summary)
with st.expander("KCB recs"):
    st.dataframe(kcb_recs_data)
with st.expander("Equity recs"):
    st.dataframe(equity_recs_data)
with st.expander("Aspire recs"):
    st.dataframe(aspire_recs_data)
with st.expander("Merged Cards"):
    st.dataframe(merged_cards)
with st.expander("New Aspire"):
    st.dataframe(newaspire)

st.success("✅ All workflows/processes (including download and all requested sheets) completed.")
