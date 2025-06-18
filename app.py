import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide", page_title="Cards Reconciliation Workbook")

st.title("Cards Reconciliation and Banking Variance Workbook")
st.markdown("""
**Instructions:**  
1. Upload the following files in their respective slots:
   - KCB Excel: _e.g. QUICK MART 11.6.2025.xlsx_
   - Equity Excel: _e.g. QUICKMART 11062025.xlsx_
   - Aspire CSV: _e.g. ZEDS_CARDS_TILLWISE_2025-06-11.csv_
   - Card Key Excel: _e.g. card_key.xlsx_

2. Click **Process and Download** to generate the workbook with all reports.
""")

def clean_ref_no(x):
    try:
        return str(int(float(x)))
    except:
        return ""

# ---- File uploads ----
kcb_file = st.file_uploader("Upload KCB Excel", type=["xlsx"])
equity_file = st.file_uploader("Upload Equity Excel", type=["xlsx"])
aspire_file = st.file_uploader("Upload Aspire CSV", type=["csv"])
key_file = st.file_uploader("Upload Card Key Excel", type=["xlsx"])

if st.button("Process and Download") and kcb_file and equity_file and aspire_file and key_file:
    try:
        # Load files
        kcb = pd.read_excel(kcb_file)
        equity = pd.read_excel(equity_file)
        aspire = pd.read_csv(aspire_file)
        key = pd.read_excel(key_file)
        card_key = key.copy()

        # Clean col names
        kcb.columns = kcb.columns.str.strip()
        equity.columns = equity.columns.str.strip()

        # KCB column renaming and cleaning
        kcb_renamed = kcb.rename(columns={
            'Card No': 'Card_Number',
            'Trans Date': 'TRANS_DATE',
            'RRN': 'R_R_N',
            'Amount': 'Purchase',
            'Comm': 'Commission',
            'NetPaid': 'Settlement_Amount',
            'Merchant': 'store'
        })
        kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -1*x if x < 0 else 0)
        kcb_renamed['Source'] = 'KCB'

        # Equity
        equity = equity.rename(columns={'Outlet_Name': 'store'})
        equity['Source'] = 'Equity'

        # Final columns
        columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
                   'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
        kcb_final = kcb_renamed[columns]
        equity_final = equity[columns]
        merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)

        # Drop Card_Number NaN/blanks
        merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
        merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']

        # Card key
        card_key.columns = card_key.columns.str.strip()
        card_key['Col_1'] = card_key['Col_1'].astype(str).str.strip()
        card_key['Col_2'] = card_key['Col_2'].astype(str).str.strip()
        merged_cards['store'] = merged_cards['store'].astype(str).str.strip()
        lookup_dict = dict(zip(card_key['Col_1'], card_key['Col_2']))
        merged_cards['branch'] = merged_cards['store'].map(lookup_dict)
        cols = list(merged_cards.columns)
        source_index = cols.index('Source')
        cols.insert(source_index + 1, cols.pop(cols.index('branch')))
        merged_cards = merged_cards[cols]

        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
        merged_cards['card_check'] = merged_cards['Card_Number'].apply(
            lambda x: x[:4] + x[-4:] if len(x.replace(" ", "").replace("*", "")) >= 8 else ''
        )
        cols = merged_cards.columns.tolist()
        if 'branch' in cols and 'card_check' in cols:
            cols.remove('card_check')
            branch_index = cols.index('branch')
            cols.insert(branch_index + 1, 'card_check')
            merged_cards = merged_cards[cols]

        # Aspire
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
        aspire = aspire[[
            'STORE_CODE', 'STORE_NAME', 'ZED_DATE', 'TILL', 'SESSION', 'RCT',
            'CUSTOMER_NAME', 'CARD_TYPE', 'CARD_NUMBER', 'card_check', 'AMOUNT', 'REF_NO', 'RCT_TRN_DATE'
        ]]
        aspire = aspire.rename(columns={'REF_NO': 'R_R_N'})
        aspire['R_R_N'] = aspire['R_R_N'].astype(str).str.strip()
        merged_cards['R_R_N'] = merged_cards['R_R_N'].astype(str).str.strip()

        # rrntable merge
        rrntable = pd.merge(
            aspire, merged_cards, on='R_R_N', how='inner', suffixes=('_aspire', '_merged')
        )

        # Unmatched branch
        missing_branch_rows = merged_cards[merged_cards['branch'].isna()]

        # card_check in merged_cards (repeated for safety)
        merged_cards['Card_Number'] = merged_cards['Card_Number'].astype(str).str.strip()
        merged_cards['card_check'] = merged_cards['Card_Number'].str[:4] + merged_cards['Card_Number'].str[-4:]
        if 'Source' in merged_cards.columns:
            cols = merged_cards.columns.tolist()
            if 'card_check' in cols:
                cols.remove('card_check')
            source_index = cols.index('Source') + 1
            cols.insert(source_index, 'card_check')
            merged_cards = merged_cards[cols]

        # Remove duplicate rows
        merged_cards = merged_cards.drop_duplicates()
        merged_cards = merged_cards[merged_cards['TID'].notna() & (merged_cards['TID'].astype(str).str.strip() != '')]

        # RRN check/val_check for aspire
        card_key = key.copy()
        merged_cards['REF_NO'] = merged_cards['R_R_N'].apply(clean_ref_no)
        aspire['REF_NO'] = aspire['R_R_N'].astype(str)
        aspire['REF_NO'] = aspire['REF_NO'].astype(str).str.lstrip('0')
        ref_to_purchase = dict(zip(merged_cards['REF_NO'], merged_cards['Purchase']))
        aspire['rrn_check'] = aspire['REF_NO'].map(ref_to_purchase).fillna(0)
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        aspire['val_check'] = aspire['AMOUNT'] - aspire['rrn_check']
        cols = list(aspire.columns)
        rrn_idx = cols.index('rrn_check')
        new_order = cols[:rrn_idx + 1] + ['val_check'] + cols[rrn_idx + 1:]
        aspire = aspire[new_order]
        aspire = aspire.loc[:, ~aspire.columns.duplicated()]

        # Cheked_rows in merged_cards
        matched_ref_nos = set(aspire['REF_NO'])
        merged_cards['Cheked_rows'] = merged_cards['REF_NO'].astype(str).apply(lambda x: 'Yes' if x in matched_ref_nos else 'No')

        # newbankmerged
        newbankmerged = merged_cards[merged_cards['Cheked_rows'].str.strip().str.upper() == 'NO'].copy()
        newbankmerged['store'] = newbankmerged['store'].astype(str)
        key['Col_1'] = key['Col_1'].astype(str)
        newbankmerged = newbankmerged.merge(
            key[['Col_1', 'Col_2']], how='left', left_on='store', right_on='Col_1'
        )
        newbankmerged = newbankmerged.rename(columns={'Col_2': 'branch'}).drop(columns=['Col_1'])
        newbankmerged['branch'] = newbankmerged['branch'].astype(str)
        newbankmerged['Purchase'] = newbankmerged['Purchase'].astype(str)
        newbankmerged['Check_Two'] = newbankmerged['branch'] + newbankmerged['Purchase']
        newbankmerged['Check_Two'] = (
            newbankmerged.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newbankmerged['Check_Two']
        newbankmerged['Check_Two'] = (
            newbankmerged['branch'].astype(str).str.strip().str.upper() +
            newbankmerged['Purchase'].astype(float).map('{:.2f}'.format)
        )
        newbankmerged['Check_Two'] = (
            newbankmerged.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newbankmerged['Check_Two']

        # newaspire
        aspire['rrn_check'] = aspire['rrn_check'].astype(float)
        newaspire = aspire[aspire['rrn_check'] <= 0].copy()
        newaspire['Check_Two'] = newaspire['STORE_NAME'].astype(str) + newaspire['AMOUNT'].map('{:.2f}'.format)
        newaspire['Check_Two'] = (
            newaspire.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newaspire['Check_Two']
        newaspire['Check_Two'] = newaspire['STORE_NAME'].astype(str).str.strip().str.upper() + \
                                 newaspire['AMOUNT'].astype(float).astype(int).astype(str)
        newbankmerged['Check_Two'] = newbankmerged['branch'].astype(str).str.strip().str.upper() + \
                                     newbankmerged['Purchase'].astype(float).astype(int).astype(str)
        available_matches = newbankmerged['Check_Two'].tolist()
        def check_and_consume(val):
            if val in available_matches:
                available_matches.remove(val)
                return 'Okay'
            else:
                return 'False'
        newaspire['Amount_check'] = newaspire['Check_Two'].apply(check_and_consume)
        newaspire['Check_Two'] = (
            newaspire.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newaspire['Check_Two']
        false_list = newaspire[newaspire['Amount_check'] == 'False'].copy()
        newbankmerged['Check_Two'] = (
            newbankmerged.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newbankmerged['Check_Two']
        aspire_available_matches = newaspire['Check_Two'].tolist()
        def check_and_consume_from_aspire(val):
            if val in aspire_available_matches:
                aspire_available_matches.remove(val)
                return 'Okay'
            else:
                return 'False'
        newbankmerged['Amount_check'] = newbankmerged['Check_Two'].apply(check_and_consume_from_aspire)

        # newmerged_cards
        if 'newmerged_cards' not in locals():
            newmerged_cards = newbankmerged[
                (newbankmerged['Cheked_rows'].str.strip().str.lower() == 'no') &
                (newbankmerged['Amount_check'].astype(str).str.strip().str.lower() == 'false')
            ].copy()
        # Standardize for matching
        newaspire['Check_Two'] = (
            newaspire['STORE_NAME'].astype(str)
            .str.replace(r'\s+', '', regex=True)
            .str.upper() +
            newaspire['AMOUNT'].astype(float).round().astype(int).astype(str)
        )
        newmerged_cards['Check_Two'] = (
            newmerged_cards['branch'].astype(str)
            .str.replace(r'\s+', '', regex=True)
            .str.upper() +
            newmerged_cards['Purchase'].astype(float).round().astype(int).astype(str)
        )
        valid_check_twos = set(newaspire['Check_Two'])
        newmerged_cards['Matchable'] = newmerged_cards['Check_Two'].isin(valid_check_twos)
        available_aspire_matches = newaspire['Check_Two'].tolist()
        def match_and_consume(val):
            if val in available_aspire_matches:
                available_aspire_matches.remove(val)
                return 'Okay'
            return 'False'
        newmerged_cards['Amount_check'] = 'False'
        newmerged_cards.loc[newmerged_cards['Matchable'], 'Amount_check'] = newmerged_cards.loc[
            newmerged_cards['Matchable'], 'Check_Two'
        ].apply(match_and_consume)
        newmerged_cards['Check_Two'] = (
            newmerged_cards.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newmerged_cards['Check_Two']
        newmerged_cards['store'] = newmerged_cards['store'].astype(str).str.strip().str.upper()
        card_key['Col_1'] = card_key['Col_1'].astype(str).str.strip().str.upper()
        newmerged_cards.drop(columns=['branch'], errors='ignore', inplace=True)
        newmerged_cards = newmerged_cards.merge(
            card_key[['Col_1', 'Col_2']],
            how='left',
            left_on='store',
            right_on='Col_1'
        )
        newmerged_cards.rename(columns={'Col_2': 'branch'}, inplace=True)
        newmerged_cards.drop(columns=['Col_1'], inplace=True)
        newmerged_cards['Check_Two'] = (
            newmerged_cards.groupby('Check_Two').cumcount() + 1
        ).astype(str) + newmerged_cards['Check_Two']
        newmerged_cards = newmerged_cards[newmerged_cards['Matchable'] != True]

        # --- Card Summary Table ---
        card_summary = aspire['STORE_NAME'].dropna().drop_duplicates().sort_values().reset_index(drop=True).to_frame(name='STORE_NAME')
        card_summary.index = card_summary.index + 1
        card_summary.reset_index(inplace=True)
        card_summary.rename(columns={'index': 'No'}, inplace=True)
        aspire['AMOUNT'] = pd.to_numeric(aspire['AMOUNT'], errors='coerce')
        aspire_sums = aspire.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
        aspire_sums = aspire_sums.rename(columns={'AMOUNT': 'Aspire_Zed'})
        card_summary = card_summary.merge(aspire_sums, on='STORE_NAME', how='left')
        card_summary['Aspire_Zed'] = card_summary['Aspire_Zed'].fillna(0)
        merged_cards['Purchase'] = pd.to_numeric(merged_cards['Purchase'], errors='coerce')
        kcb_grouped = (
            merged_cards[merged_cards['Source'] == 'KCB']
            .groupby('branch')['Purchase']
            .sum().reset_index()
            .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'kcb_paid'})
        )
        card_summary = card_summary.merge(kcb_grouped, on='STORE_NAME', how='left')
        card_summary['kcb_paid'] = card_summary['kcb_paid'].fillna(0)
        equity_grouped = (
            merged_cards[merged_cards['Source'] == 'Equity']
            .groupby('branch')['Purchase']
            .sum().reset_index()
            .rename(columns={'branch': 'STORE_NAME', 'Purchase': 'equity_paid'})
        )
        card_summary = card_summary.merge(equity_grouped, on='STORE_NAME', how='left')
        card_summary['equity_paid'] = card_summary['equity_paid'].fillna(0)
        cols = list(card_summary.columns)
        if 'kcb_paid' in cols and 'equity_paid' in cols:
            kcb_index = cols.index('kcb_paid')
            cols.insert(kcb_index + 1, cols.pop(cols.index('equity_paid')))
            card_summary = card_summary[cols]

        # Totals row
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid']
        for col in numeric_cols:
            card_summary[col] = card_summary[col].astype(float)
        card_summary['Gross_Banking'] = card_summary['kcb_paid'] + card_summary['equity_paid']
        card_summary['Variance'] = card_summary['Gross_Banking'] - card_summary['Aspire_Zed']

        # =========================
        # STRICT CARD RECS STEPS
        # =========================

        # === KCB unmatched strict steps ===
        # ✅ Step 1: Filter relevant KCB unmatched records
        kcb_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'KCB') &
            (newmerged_cards['Amount_check'].astype(str).str.strip().str.lower() == 'false')
        ]
        # ✅ Step 2: Ensure 'Purchase' is numeric
        kcb_recs_data['Purchase'] = pd.to_numeric(kcb_recs_data['Purchase'], errors='coerce')
        # ✅ Drop rows where Purchase is NaN (can't sum invalids)
        kcb_recs_data = kcb_recs_data.dropna(subset=['Purchase'])
        # ✅ Step 3: Group by branch and sum Purchase
        kcb_recs_grouped = kcb_recs_data.groupby('branch')['Purchase'].sum().reset_index()
        kcb_recs_grouped.columns = ['STORE_NAME', 'kcb_recs']
        # ✅ Step 4: Drop existing kcb_recs if present
        if 'kcb_recs' in card_summary.columns:
            card_summary.drop(columns=['kcb_recs'], inplace=True)
        # ✅ Step 5: Merge with card_summary
        card_summary = card_summary.merge(kcb_recs_grouped, on='STORE_NAME', how='left')
        card_summary['kcb_recs'] = card_summary['kcb_recs'].fillna(0)
        # ✅ Step 6: Remove old TOTAL row
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        # ✅ Step 7: Ensure numeric columns are properly typed
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking', 'Variance', 'kcb_recs']
        card_summary[numeric_cols] = card_summary[numeric_cols].apply(pd.to_numeric, errors='coerce')
        # ✅ Step 8: Compute totals
        totals = card_summary[numeric_cols].sum()
        # ✅ Step 9: Build TOTAL row
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs']
        }])
        # ✅ Step 10: Append TOTAL row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        # ✅ Step 11: Display final result (Streamlit will export later)

        # === Equity unmatched strict steps ===
        equity_recs_data = newmerged_cards[
            (newmerged_cards['Source'].str.upper() == 'EQUITY') &
            (newmerged_cards['Amount_check'] == 'False')
        ].copy()
        # ✅ Step 1: Ensure Purchase column is numeric
        equity_recs_data['Purchase'] = pd.to_numeric(equity_recs_data['Purchase'], errors='coerce')
        # ✅ Step 2: Drop any rows with NaN in Purchase
        equity_recs_data = equity_recs_data.dropna(subset=['Purchase'])
        # ✅ Step 3: Group by branch and sum Purchase
        equity_recs_grouped = equity_recs_data.groupby('branch')['Purchase'].sum().reset_index()
        equity_recs_grouped.columns = ['STORE_NAME', 'Equity_recs']
        # ✅ Step 4: Drop old Equity_recs if exists
        if 'Equity_recs' in card_summary.columns:
            card_summary.drop(columns=['Equity_recs'], inplace=True)
        # ✅ Step 5: Merge into card_summary
        card_summary = card_summary.merge(equity_recs_grouped, on='STORE_NAME', how='left')
        card_summary['Equity_recs'] = card_summary['Equity_recs'].fillna(0)
        # ✅ Step 6: Remove old TOTAL row
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        # ✅ Step 7: Define all numeric columns
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                        'Variance', 'kcb_recs', 'Equity_recs']
        # ✅ Step 8: Convert all numeric columns to numeric types
        card_summary[numeric_cols] = card_summary[numeric_cols].apply(pd.to_numeric, errors='coerce')
        # ✅ Step 9: Compute totals
        totals = card_summary[numeric_cols].sum()
        # ✅ Step 10: Build TOTAL row
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs'],
            'Equity_recs': totals['Equity_recs']
        }])
        # ✅ Step 11: Append TOTAL row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)
        # ✅ Step 12: Display result (Streamlit will export later)

        # === Aspire unmatched strict steps ===
        aspire_recs_data = newaspire[
            newaspire['Amount_check'].astype(str).str.strip().str.lower() == 'false'
        ]
        # ✅ Step 2: Ensure AMOUNT is numeric
        aspire_recs_data['AMOUNT'] = pd.to_numeric(aspire_recs_data['AMOUNT'], errors='coerce')
        # ✅ Step 3: Drop rows with NaN in AMOUNT before summing
        aspire_recs_data = aspire_recs_data.dropna(subset=['AMOUNT'])
        # ✅ Step 4: Group by STORE_NAME and sum AMOUNT
        aspire_recs_grouped = aspire_recs_data.groupby('STORE_NAME')['AMOUNT'].sum().reset_index()
        aspire_recs_grouped.columns = ['STORE_NAME', 'Asp_Recs']
        # ✅ Step 5: Drop old Asp_Recs if exists
        if 'Asp_Recs' in card_summary.columns:
            card_summary.drop(columns=['Asp_Recs'], inplace=True)
        # ✅ Step 6: Merge into card_summary
        card_summary = card_summary.merge(aspire_recs_grouped, on='STORE_NAME', how='left')
        card_summary['Asp_Recs'] = card_summary['Asp_Recs'].fillna(0)
        # ✅ Step 7: Remove old TOTAL row
        card_summary = card_summary[card_summary['STORE_NAME'] != 'TOTAL']
        # ✅ Step 8: Recalculate TOTAL row including Asp_Recs
        numeric_cols = ['Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                        'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']
        card_summary[numeric_cols] = card_summary[numeric_cols].apply(pd.to_numeric, errors='coerce')
        totals = card_summary[numeric_cols].sum()
        # ✅ Step 9: Build TOTAL row
        total_row = pd.DataFrame([{
            'No': '',
            'STORE_NAME': 'TOTAL',
            'Aspire_Zed': totals['Aspire_Zed'],
            'kcb_paid': totals['kcb_paid'],
            'equity_paid': totals['equity_paid'],
            'Gross_Banking': totals['Gross_Banking'],
            'Variance': totals['Variance'],
            'kcb_recs': totals['kcb_recs'],
            'Equity_recs': totals['Equity_recs'],
            'Asp_Recs': totals['Asp_Recs']
        }])
        # ✅ Step 10: Append total row
        card_summary = pd.concat([card_summary, total_row], ignore_index=True)

        # ======= Net variance & Reordering columns =======
        for col in ['Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs']:
            if col not in card_summary.columns:
                card_summary[col] = 0
        card_summary['Net_variance'] = (
            card_summary['Variance']
            - card_summary['kcb_recs']
            - card_summary['Equity_recs']
            + card_summary['Asp_Recs']
        )
        final_cols = ['No', 'STORE_NAME', 'Aspire_Zed', 'kcb_paid', 'equity_paid', 'Gross_Banking',
                      'Variance', 'kcb_recs', 'Equity_recs', 'Asp_Recs', 'Net_variance']
        card_summary = card_summary[[col for col in final_cols if col in card_summary.columns]]

        # -------------- Final Export --------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            card_summary.to_excel(writer, sheet_name='card_summary', index=False)
            kcb_recs_data.to_excel(writer, sheet_name='kcb_recs', index=False)
            equity_recs_data.to_excel(writer, sheet_name='Equity_recs', index=False)
            aspire_recs_data.to_excel(writer, sheet_name='Asp_Recs', index=False)
            merged_cards.to_excel(writer, sheet_name='merged_cards', index=False)
            newaspire.to_excel(writer, sheet_name='newaspire', index=False)
        output.seek(0)
        st.success("✅ All reports exported successfully. Download below.")
        st.download_button(
            label="Download Reconciliation_Report.xlsx",
            data=output,
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error in processing: {e}")
