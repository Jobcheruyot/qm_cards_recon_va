import pandas as pd
import streamlit as st

st.set_page_config(layout="wide", page_title="Card Reconciliation Report", initial_sidebar_state="expanded")

st.title("Card Reconciliation Report")
st.write(
    "Upload your KCB, Equity, Aspire, and Card Key files in the sidebar. "
    "All processing is in-memory. Download ONE Excel workbook with all reconciliation sheets."
)

# ---------- SIDEBAR: FILE UPLOADS ----------
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

# ---------- LOAD FILES ----------
kcb = pd.read_excel(kcb_file)
equity = pd.read_excel(equity_file)
aspire = pd.read_csv(aspire_file)
key = pd.read_excel(cardkey_file)

# ---------- CLEAN COLUMNS ----------
kcb.columns = kcb.columns.str.strip()
equity.columns = equity.columns.str.strip()
key.columns = key.columns.str.strip()

# ---------- KCB PROCESSING ----------
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

# ---------- EQUITY PROCESSING ----------
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

# ---------- MAP BRANCHES VIA CARD KEY ----------
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

# ---------- ADD CARD CHECK COLUMN ----------
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

# ---------- ASPIRE PROCESSING ----------
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

# ---------- RRN MERGE (ASPIRE x BANK) ----------
rrntable = pd.merge(
    aspire,
    merged_cards,
    on='R_R_N',
    how='inner',
    suffixes=('_aspire', '_merged')
)

# ----------- Identify unmatched branch rows for key export -----------
missing_branch_rows = merged_cards[merged_cards['branch'].isna()]

# ---------- CARD SUMMARY ----------
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

# Add TOTAL row
numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance']
totals = card_summary[numeric_cols].sum()
total_row = pd.DataFrame([{
    'No': '',
    'STORE_NAME': 'TOTAL',
    'Aspire_Zed': totals['Aspire_Zed'],
    'kcb_paid': totals['kcb_paid'],
    'equity_paid': totals['equity_paid'],
    'Gross_Banking': totals['Gross_Banking'],
    'Variance': totals['Variance']
}])
card_summary = pd.concat([card_summary, total_row], ignore_index=True)

# ========== SINGLE DOWNLOAD BUTTON FOR FULL EXCEL REPORT ==========
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
        aspire=aspire,
        merged_cards=merged_cards,
        rrntable=rrntable,
        missing_branch_rows=missing_branch_rows
    ),
    file_name="Reconciliation_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Optional: Data Previews for transparency
with st.expander("Card Summary (Final Preview)"):
    st.dataframe(card_summary)
with st.expander("Aspire (Filtered/Aligned) Preview"):
    st.dataframe(aspire)
with st.expander("Merged Cards Preview"):
    st.dataframe(merged_cards)
with st.expander("RRN Table (Aspire x Bank Merge)"):
    st.dataframe(rrntable)
with st.expander("Missing branch/Key rows"):
    st.dataframe(missing_branch_rows)

st.success("✅ All workflows/processes complete. One download with all reconciliation data!")
