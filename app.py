import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import datetime
import re

# Page configuration
st.set_page_config(
    page_title="Card Transaction Reconciliation",
    page_icon="💳",
    layout="wide"
)

# Custom CSS for styling
st.markdown("""
    <style>
    .main {background-color: #f5f5f5;}
    .report-title {font-size: 24px; color: #2c3e50; text-align: center;}
    .sidebar .sidebar-content {background-color: #2c3e50;}
    .stButton>button {background-color: #3498db; color: white;}
    .stDownloadButton>button {background-color: #2ecc71; color: white;}
    .stAlert {border-left: 5px solid #3498db;}
    .metric-card {background: white; border-radius: 10px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);}
    .bank-metrics {display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 20px;}
    </style>
    """, unsafe_allow_html=True)

# App title
st.title("💳 Card Transaction Reconciliation")
st.markdown("Upload bank statements and get reconciled reports")

# Sidebar for file uploads
with st.sidebar:
    st.header("Upload Files")
    
    # File uploaders
    kcb_file = st.file_uploader("KCB Statement (Excel)", type=["xlsx"])
    equity_file = st.file_uploader("Equity Statement (Excel)", type=["xlsx"])
    coop_file = st.file_uploader("Co-op Statement (Excel)", type=["xlsx", "xls"])
    aspire_file = st.file_uploader("Aspire Statement (CSV)", type=["csv"])
    key_file = st.file_uploader("Branch Key (Excel)", type=["xlsx"])
    
    # Date selector
    report_date = st.date_input("Report Date", datetime.date.today())
    
    # Process button
    process_btn = st.button("Process Statements")

# Main processing function
def process_statements():
    # Initialize empty DataFrames
    dfs = {
        'KCB': pd.DataFrame(),
        'Equity': pd.DataFrame(),
        'Co-op': pd.DataFrame(),
        'Aspire': pd.DataFrame()
    }
    key = pd.DataFrame()
    
    # Load uploaded files
    try:
        if kcb_file:
            dfs['KCB'] = pd.read_excel(kcb_file)
        if equity_file:
            dfs['Equity'] = pd.read_excel(equity_file)
        if coop_file:
            dfs['Co-op'] = pd.read_excel(coop_file, skiprows=6)
        if aspire_file:
            dfs['Aspire'] = pd.read_csv(aspire_file)
        if key_file:
            key = pd.read_excel(key_file)
    except Exception as e:
        st.error(f"Error loading files: {str(e)}")
        return None, None
    
    # Process each bank's data
    try:
        # Process KCB data
        if not dfs['KCB'].empty:
            dfs['KCB'].columns = dfs['KCB'].columns.str.strip()
            dfs['KCB']['Amount'] = pd.to_numeric(dfs['KCB']['Amount'], errors='coerce')
            dfs['KCB'] = dfs['KCB'].drop_duplicates(subset=['RRN', 'Amount'], keep='first')
            dfs['KCB']['Source'] = 'KCB'
        
        # Process Co-op data
        if not dfs['Co-op'].empty:
            dfs['Co-op'].columns = dfs['Co-op'].columns.str.strip()
            dfs['Co-op']['BANK COMM'] = pd.to_numeric(dfs['Co-op']['BANK COMM'], errors='coerce')
            dfs['Co-op'] = dfs['Co-op'].sort_values(by='BANK COMM', na_position='first')
            dfs['Co-op'] = dfs['Co-op'].drop_duplicates(subset='RRN CODE', keep='first')
            dfs['Co-op']['Source'] = 'Co-op'
            dfs['Co-op'] = dfs['Co-op'].dropna(subset=["TRANSACTION DATE"]).reset_index(drop=True)
        
        # Process Equity data
        if not dfs['Equity'].empty:
            dfs['Equity'].columns = dfs['Equity'].columns.str.strip()
            dfs['Equity']['Commission'] = pd.to_numeric(dfs['Equity']['Commission'], errors='coerce')
            dfs['Equity'] = dfs['Equity'].sort_values(by='Commission', na_position='first')
            dfs['Equity'] = dfs['Equity'].drop_duplicates(subset='R_R_N', keep='first')
            dfs['Equity']['Source'] = 'Equity'
        
        # Process Aspire data (if needed)
        if not dfs['Aspire'].empty:
            dfs['Aspire']['Source'] = 'Aspire'
        
        # Merge KCB and Equity
        merged_cards = pd.DataFrame()
        if not dfs['KCB'].empty and not dfs['Equity'].empty:
            # Rename columns for consistency
            kcb_renamed = dfs['KCB'].rename(columns={
                'Card No': 'Card_Number',
                'Trans Date': 'TRANS_DATE',
                'RRN': 'R_R_N',
                'Amount': 'Purchase',
                'Comm': 'Commission',
                'NetPaid': 'Settlement_Amount',
                'Merchant': 'store'
            })
            kcb_renamed['Cash_Back'] = kcb_renamed['Purchase'].apply(lambda x: -1 * x if x < 0 else 0)
            
            equity_renamed = dfs['Equity'].rename(columns={
                'Outlet_Name': 'store',
                'Trans_Amount': 'Purchase',
                'Settlement_Amount': 'Settlement_Amount',
                'R_R_N': 'R_R_N'
            })
            
            # Select common columns
            columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
                      'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
            
            # Merge
            kcb_final = kcb_renamed[columns]
            equity_final = equity_renamed[columns]
            merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)
            
            # Clean merged data
            merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
            merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']
            
            # Standardize card numbers
            def standardize_card_number(card_num):
                if pd.isna(card_num):
                    return card_num
                card_str = str(card_num)
                # Remove all non-digit characters
                digits = re.sub(r'\D', '', card_str)
                # Mask middle digits if long enough
                if len(digits) >= 12:
                    return f"{digits[:6]}******{digits[-4:]}"
                return card_str
            
            merged_cards['Card_Number'] = merged_cards['Card_Number'].apply(standardize_card_number)
            
            # Add branch information if key file is provided
            if not key.empty:
                # Create branch mapping dictionary
                branch_mapping = dict(zip(key['Col_1'], key['Col_2']))
                
                # Function to extract branch from store name
                def get_branch(store_name):
                    store_name = str(store_name).upper()
                    for key in branch_mapping:
                        if str(key).upper() in store_name:
                            return branch_mapping[key]
                    # Try to extract branch from KCB merchant format
                    if "QUICK MART" in store_name and "TILL" not in store_name:
                        parts = store_name.split(",")
                        if len(parts) >= 2:
                            return parts[0].split("QUICK MART")[-1].strip("- ").strip()
                    return "UNKNOWN"
                
                merged_cards['branch'] = merged_cards['store'].apply(get_branch)
        
        return merged_cards, dfs
    
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        return None, None

# Main content area
if process_btn:
    if not (kcb_file or equity_file or coop_file or aspire_file):
        st.warning("Please upload at least one bank statement")
    else:
        with st.spinner("Processing statements..."):
            merged_cards, dfs = process_statements()
            
            if merged_cards is not None:
                st.success("Processing completed!")
                
                # Display comprehensive statistics
                st.subheader("Comprehensive Statistics")
                
                # Create metrics for each bank
                st.markdown("### Transaction Summary by Bank")
                
                # Calculate metrics for each available bank
                bank_metrics = {}
                for bank in ['KCB', 'Equity', 'Co-op', 'Aspire']:
                    if not dfs[bank].empty:
                        if bank in ['KCB', 'Equity']:
                            amount_col = 'Amount' if bank == 'KCB' else 'Purchase'
                            comm_col = 'Comm' if bank == 'KCB' else 'Commission'
                            count = len(dfs[bank])
                            total = dfs[bank][amount_col].sum()
                            commission = dfs[bank][comm_col].sum()
                        elif bank == 'Co-op':
                            count = len(dfs[bank])
                            total = dfs[bank]['TRANSACTION AMOUNT'].sum()
                            commission = dfs[bank]['BANK COMM'].sum()
                        else:  # Aspire
                            count = len(dfs[bank])
                            total = dfs[bank].get('Amount', pd.Series([0])).sum()
                            commission = 0  # Adjust based on actual Aspire data
                        
                        bank_metrics[bank] = {
                            'Transactions': count,
                            'Total Amount': total,
                            'Total Commission': commission
                        }
                
                # Display metrics in cards
                cols = st.columns(len(bank_metrics))
                for idx, (bank, metrics) in enumerate(bank_metrics.items()):
                    with cols[idx]:
                        st.markdown(f"<div class='metric-card'><h3>{bank}</h3>"
                                   f"<p>Transactions: {metrics['Transactions']:,}</p>"
                                   f"<p>Amount: KES {metrics['Total Amount']:,.2f}</p>"
                                   f"<p>Commission: KES {metrics['Total Commission']:,.2f}</p></div>", 
                                   unsafe_allow_html=True)
                
                # Show merged data statistics if available
                if not merged_cards.empty:
                    st.markdown("### Merged Data Summary")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Transactions", len(merged_cards))
                    with col2:
                        st.metric("Total Amount", f"KES {merged_cards['Purchase'].sum():,.2f}")
                    with col3:
                        st.metric("Total Commission", f"KES {merged_cards['Commission'].sum():,.2f}")
                    
                    # Show source distribution
                    st.write("#### Transactions by Bank")
                    source_counts = merged_cards['Source'].value_counts()
                    st.bar_chart(source_counts)
                    
                    # Show branch distribution if available
                    if 'branch' in merged_cards.columns:
                        st.write("#### Transactions by Branch")
                        branch_counts = merged_cards['branch'].value_counts()
                        st.bar_chart(branch_counts)
                
                # Show data previews
                st.subheader("Data Previews")
                
                tab1, tab2, tab3, tab4 = st.tabs(["Merged Data", "KCB", "Equity", "Co-op"])
                
                with tab1:
                    if not merged_cards.empty:
                        st.dataframe(merged_cards.head())
                    else:
                        st.info("No merged data available")
                
                with tab2:
                    if not dfs['KCB'].empty:
                        st.dataframe(dfs['KCB'].head())
                    else:
                        st.info("No KCB data available")
                
                with tab3:
                    if not dfs['Equity'].empty:
                        st.dataframe(dfs['Equity'].head())
                    else:
                        st.info("No Equity data available")
                
                with tab4:
                    if not dfs['Co-op'].empty:
                        st.dataframe(dfs['Co-op'].head())
                    else:
                        st.info("No Co-op data available")
                
                # Download buttons
                st.subheader("Download Reports")
                
                # Create Excel file with multiple sheets
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    if not merged_cards.empty:
                        # Format merged data to match reconciliation report
                        report_df = merged_cards.copy()
                        report_df['Transaction_Date'] = pd.to_datetime(report_df['TRANS_DATE']).dt.strftime('%Y-%m-%d %H:%M:%S')
                        report_df = report_df[[
                            'Transaction_Date', 'branch', 'Card_Number', 'Purchase', 
                            'Commission', 'Settlement_Amount', 'Source', 'R_R_N', 'TID'
                        ]]
                        report_df.to_excel(writer, sheet_name='Reconciled_Transactions', index=False)
                    
                    # Add individual bank sheets
                    for bank, df in dfs.items():
                        if not df.empty:
                            df.to_excel(writer, sheet_name=f'{bank}_Raw_Data', index=False)
                
                # Create download link
                b64 = base64.b64encode(output.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Reconciliation_Report_{report_date}.xlsx">Download Full Report</a>'
                st.markdown(href, unsafe_allow_html=True)

# Instructions section
with st.expander("📌 Instructions"):
    st.markdown("""
    ### How to use this tool:
    1. **Upload Files** in the sidebar:
       - KCB Statement (Excel)
       - Equity Statement (Excel)
       - Co-op Statement (Excel)
       - Aspire Statement (CSV)
       - Branch Key (Excel) - Optional but recommended for branch mapping
    
    2. Select the **report date**
    
    3. Click **"Process Statements"** button
    
    4. View and download the reconciled reports
    
    ### Expected File Formats:
    - **KCB**: Excel with columns: Card No, Trans Date, RRN, Amount, Comm, NetPaid, Merchant
    - **Equity**: Excel with columns: Outlet_Name, Card_Number, TRANS_DATE, R_R_N, Purchase, Commission, Settlement_Amount
    - **Co-op**: Excel file with transaction data starting from row 7
    - **Aspire**: CSV file with transaction data
    - **Branch Key**: Excel with two columns mapping store names to branches
    """)

# About section
st.sidebar.markdown("---")
st.sidebar.markdown("""
**About This App**  
Card transaction reconciliation tool that:
- Processes multiple bank statements
- Identifies duplicate transactions
- Provides branch-level reporting
- Generates reconciliation reports

Version 2.0  
""")
