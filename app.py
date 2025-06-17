
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Card Reconciliation Report", layout="wide")
st.title("📊 Card Reconciliation Report")

st.markdown("Upload your KCB, Equity, Aspire, and Card Key files in the sidebar. All processing is in-memory. One Excel workbook with all reconciliation sheets will be generated.")

# Upload widgets
kcb_file = st.sidebar.file_uploader("Upload KCB Excel File", type=["xlsx"])
equity_file = st.sidebar.file_uploader("Upload Equity Excel File", type=["xlsx"])
aspire_file = st.sidebar.file_uploader("Upload Aspire CSV File", type=["csv"])
key_file = st.sidebar.file_uploader("Upload Card Key Excel File", type=["xlsx"])

if st.sidebar.button("🚀 Run Reconciliation") and all([kcb_file, equity_file, aspire_file, key_file]):
    st.success("Processing...")

    # Read inputs
    kcb = pd.read_excel(kcb_file)
    equity = pd.read_excel(equity_file)
    aspire = pd.read_csv(aspire_file)
    key = pd.read_excel(key_file)

    # === Original logic from cards.py goes here ===
with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:

    card_summary.to_excel(writer, sheet_name='card_summary', index=False)

    aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)

    equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)

    kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)

    merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)

    aspire.to_excel(writer, sheet_name='aspire', index=False)



print("✅ All reports exported successfully to:", filename)



# ------------------ Auto-Download in Colab ------------------



files.download(filename)



!pip install xlsxwriter



import pandas as pd



# ------------------ Prepare Sheets ------------------



# Asp_Recs (Sheet 2)

aspire_recs_data = newaspire[newaspire['Amount_check'] == 'False'].copy()



# Equity_recs (Sheet 3)

equity_recs_data = newmerged_cards[

    (newmerged_cards['Source'].str.upper() == 'EQUITY') &

    (newmerged_cards['Amount_check'] == 'False')

].copy()



# KCB_recs (New Sheet)

kcb_recs_data = newmerged_cards[

    (newmerged_cards['Source'].str.upper() == 'KCB') &

    (newmerged_cards['Amount_check'] == 'False')

].copy()



# Clean merged_cards (Sheet 5)

merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str)

merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]

if 'Source' in merged_cards.columns:

    cols = merged_cards.columns.tolist()

    if 'card_check' in cols:

        cols.remove('card_check')

    source_index = cols.index('Source') + 1

    cols.insert(source_index, 'card_check')

    merged_cards = merged_cards[cols]

merged_cards = merged_cards.drop_duplicates()

merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]



# Clean aspire (Sheet 6)

aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')

aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']

cols = list(aspire.columns)

rrn_idx = cols.index('rrn_check')

new_order = cols[:rrn_idx + 1] + ['val_check'] + cols[rrn_idx + 1:-1] + [cols[-1]]

aspire = aspire[new_order]

aspire = aspire.loc[:, ~aspire.columns.duplicated()]



# ------------------ Recalculate Total Row ------------------



# Remove old TOTAL row

card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']



# Numeric columns to total

numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',

                'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']

totals = card_summary[numeric_cols].sum()



# Create TOTAL row

total_row = pd.DataFrame([{

    'No': '',

    'STORE_NAME': 'TOTAL',

    'Aspire_Zed': totals.get('Aspire_Zed', 0),

    'kcb_paid': totals.get('kcb_paid', 0),

    'equity_paid': totals.get('equity_paid', 0),

    'Gross_Banking': totals.get('Gross_Banking', 0),

    'Variance': totals.get('Variance', 0),

    'kcb_recs': totals.get('kcb_recs', 0),

    'Equity_recs': totals.get('Equity_recs', 0),

    'Asp_Recs': totals.get('Asp_Recs', 0)

}])



# Append to card_summary

card_summary = pd.concat([card_summary, total_row], ignore_index=True)



# ------------------ Export to Excel ------------------



with pd.ExcelWriter("Reconciliation_Report.xlsx", engine='xlsxwriter') as writer:

    card_summary.to_excel(writer, sheet_name='card_summary', index=False)

    aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)

    equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)

    kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)

    merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)

    aspire.to_excel(writer, sheet_name='aspire', index=False)



print("✅ All sheets exported to 'Reconciliation_Report.xlsx'")



# Temporarily remove row display limit

pd.set_option('display.max_rows', None)



# Display full DataFrame

display(card_summary)



# (Optional) Reset display limit afterward

    # Save final output to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        card_summary.to_excel(writer, sheet_name='card_summary', index=False)
        kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)
        equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)
        aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)
        merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
        newaspire.to_excel(writer, sheet_name='newaspire', index=False)
    output.seek(0)

    st.download_button(
        label="📥 Download Reconciliation_Report.xlsx",
        data=output,
        file_name="Reconciliation_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
