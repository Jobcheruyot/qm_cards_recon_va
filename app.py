import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import base64
import datetime

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

# Main content area
if process_btn:
    if not (kcb_file or equity_file or coop_file or aspire_file):
        st.warning("Please upload at least one bank statement")
    else:
        with st.spinner("Processing statements..."):
            try:
                # Initialize empty DataFrames
                kcb = pd.DataFrame()
                equity = pd.DataFrame()
                coop = pd.DataFrame()
                aspire = pd.DataFrame()
                key = pd.DataFrame()
                
                # Load uploaded files
                if kcb_file:
                    kcb = pd.read_excel(kcb_file)
                if equity_file:
                    equity = pd.read_excel(equity_file)
                if coop_file:
                    coop = pd.read_excel(coop_file, skiprows=6)
                if aspire_file:
                    aspire = pd.read_csv(aspire_file)
                if key_file:
                    key = pd.read_excel(key_file)
                
                # Processing logic (similar to your notebook)
                # ------------------------------------------
                
                # Process KCB data
                if not kcb.empty:
                    kcb.columns = kcb.columns.str.strip()
                    kcb['Amount'] = pd.to_numeric(kcb['Amount'], errors='coerce')
                    kcb = kcb.drop_duplicates(subset=['RRN', 'Amount'], keep='first')
                    kcb['Source'] = 'KCB'
                
                # Process Co-op data
                if not coop.empty:
                    coop.columns = coop.columns.str.strip()
                    coop['BANK COMM'] = pd.to_numeric(coop['BANK COMM'], errors='coerce')
                    coop = coop.sort_values(by='BANK COMM', na_position='first')
                    coop = coop.drop_duplicates(subset='RRN CODE', keep='first')
                    coop['Source'] = 'coop'
                    coop = coop.dropna(subset=["TRANSACTION DATE"]).reset_index(drop=True)
                
                # Process Equity data
                if not equity.empty:
                    equity.columns = equity.columns.str.strip()
                    equity['Commission'] = pd.to_numeric(equity['Commission'], errors='coerce')
                    equity = equity.sort_values(by='Commission', na_position='first')
                    equity = equity.drop_duplicates(subset='R_R_N', keep='first')
                    equity['Source'] = 'Equity'
                
                # Merge KCB and Equity
                merged_cards = pd.DataFrame()
                if not kcb.empty and not equity.empty:
                    # Rename columns for consistency
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
                    
                    equity = equity.rename(columns={'Outlet_Name': 'store'})
                    
                    # Select common columns
                    columns = ['TID', 'store', 'Card_Number', 'TRANS_DATE', 'R_R_N',
                              'Purchase', 'Commission', 'Settlement_Amount', 'Cash_Back', 'Source']
                    
                    # Merge
                    kcb_final = kcb_renamed[columns]
                    equity_final = equity[columns]
                    merged_cards = pd.concat([kcb_final, equity_final], ignore_index=True)
                    
                    # Clean merged data
                    merged_cards = merged_cards[merged_cards['Card_Number'].notna()]
                    merged_cards = merged_cards[merged_cards['Card_Number'].astype(str).str.strip() != '']
                
                # Add branch information if key file is provided
                if not key.empty and not merged_cards.empty:
                    # Create branch mapping dictionary
                    branch_mapping = dict(zip(key['Col_1'], key['Col_2']))
                    
                    # Function to extract branch from store name
                    def get_branch(store_name):
                        store_name = str(store_name).upper()
                        for key in branch_mapping:
                            if str(key).upper() in store_name:
                                return branch_mapping[key]
                        return "UNKNOWN"
                    
                    merged_cards['branch'] = merged_cards['store'].apply(get_branch)
                
                # ------------------------------------------
                
                # Display results
                st.success("Processing completed!")
                
                # Show summary statistics
                st.subheader("Summary Statistics")
                
                if not merged_cards.empty:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Transactions", len(merged_cards))
                    with col2:
                        st.metric("Total Amount", f"KES {merged_cards['Purchase'].sum():,.2f}")
                    with col3:
                        st.metric("Total Commission", f"KES {merged_cards['Commission'].sum():,.2f}")
                    
                    # Show source distribution
                    st.write("### Transactions by Bank")
                    source_counts = merged_cards['Source'].value_counts()
                    st.bar_chart(source_counts)
                
                # Show data preview
                st.subheader("Processed Data Preview")
                if not merged_cards.empty:
                    st.dataframe(merged_cards.head())
                
                # Download buttons
                st.subheader("Download Reports")
                
                if not merged_cards.empty:
                    # Convert DataFrame to Excel
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        merged_cards.to_excel(writer, sheet_name='Reconciled_Transactions', index=False)
                        if not coop.empty:
                            coop.to_excel(writer, sheet_name='Coop_Transactions', index=False)
                    
                    # Create download link
                    b64 = base64.b64encode(output.getvalue()).decode()
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="Reconciled_Transactions_{report_date}.xlsx">Download Full Report</a>'
                    st.markdown(href, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

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
    - **KCB/Equity**: Excel files with standard transaction columns
    - **Co-op**: Excel file with transaction data starting from row 7
    - **Aspire**: CSV file with transaction data
    - **Branch Key**: Excel file with two columns mapping store names to branches
    """)

# About section
st.sidebar.markdown("---")
st.sidebar.markdown("""
**About This App**  
A tool for reconciling card transactions  
from multiple bank statements.  

Developed by [Your Name]  
Version 1.0  
""")
