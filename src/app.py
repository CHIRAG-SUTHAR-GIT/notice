"""
DataLens for Cyber Cell - Web Application.

A Streamlit-based web application for law enforcement cybercrime departments
to analyze and consolidate transaction data from Excel files.

Last Updated: 2026-04-27 - Account report now uses distinct Account Numbers
"""

import sys
import os

# Ensure the project root is on sys.path so `import src.xyz` works when
# Streamlit runs `src/app.py` as a script. This makes imports stable whether
# you run `python -m streamlit run src/app.py` from the repo root or from
# another working directory.
_THIS_FILE = os.path.abspath(__file__)
_PROJECT_ROOT = os.path.dirname(os.path.dirname(_THIS_FILE))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from typing import Optional, List, Dict, Any

from src.upload_service import UploadService
from src.column_detector import ColumnDetector
from src.data_processor import DataProcessor
from src.validation_engine import ValidationEngine
from src.aggregation_engine import AggregationEngine
from src.report_generator import ReportGenerator
from src.dashboard import Dashboard
from src.session_manager import SessionManager
from src.models import ColumnMapping, AggregatedAccount, ProcessingStats, ValidationResult
from src.district_data import render_district_download_page
from src.merge_files import render_merge_files_page
from src.excel_merger import render_excel_merger_page
from src.call_notice_data_merge import render_call_notice_merge_page
from src.transaction_matcher import render_transaction_matcher_page
from src.disputed_amount_matcher import render_disputed_amount_matcher_page
from src.money_transfer_dispute import render_money_transfer_dispute_page
from src.ack_bank_consolidator import render_ack_bank_consolidator_page
from src.database_service import DatabaseService
from src.report_service import ReportService
from src.districtwise import render_districtwise_page
from src.non_gujarat_filter import render_non_gujarat_filter_page
from src.column_selector import render_column_selector_page
from src.csv_fixer import render_csv_fixer_page
from src.drop_call_finder import render_drop_call_finder_page
from src.mo_finder import render_mo_finder_page
from src.amount_matcher import render_amount_matcher_page
from src.bank_ack_pivot import render_bank_ack_pivot_page
from src.ack_list_pivot import render_ack_list_pivot_page
from src.filter_by_entry_count import render_filter_by_entry_count_page
from src.filter_by_unique_ack import render_filter_by_unique_ack_page
from src.bulk_mysql_import import render_bulk_mysql_import_page
from src.mysql_database_viewer import render_mysql_database_viewer_page
from src.ai_sql_assistant import render_ai_sql_assistant_page
from src.automated_workflow import render_automated_workflow_page
from src.gujarat_account_formatter import render_gujarat_account_formatter_page
from src.kyc_notice_mailer import render_kyc_notice_mailer_page
from src.column_transfer import render_column_transfer_page
from src.ui_styling import apply_custom_css, render_page_header_with_info
from src.smart_district_split import render_smart_district_split_page
from src.ifsc_pincode_district_split import render_ifsc_pincode_district_split_page
from src.daily_report_district_split import render_daily_report_district_split_page
from src.distinct_account_pivot import render_distinct_account_pivot_page
from src.investment_fraud_top_accounts import render_investment_fraud_top_accounts_page
from src.top_10_suspect_accounts import render_top_10_suspect_accounts_page
from src.e_fir_report import render_efir_report_page

# Page configuration
st.set_page_config(
    page_title="DataLens for Cyber Cell",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom styling
apply_custom_css()

# Initialize services (cached to avoid recreation)
@st.cache_resource
def get_services():
    """Initialize and cache service instances."""
    return {
        'upload_service': UploadService(),
        'column_detector': ColumnDetector(),
        'data_processor': DataProcessor(),
        'validation_engine': ValidationEngine(),
        'aggregation_engine': AggregationEngine(),
        'report_generator': ReportGenerator(),
        'dashboard': Dashboard(),
        'session_manager': SessionManager()
    }


def init_session_state():
    """Initialize session state variables."""
    if 'current_page' not in st.session_state:
        st.session_state.current_page = 'upload'
    if 'uploaded_df' not in st.session_state:
        st.session_state.uploaded_df = None
    if 'filename' not in st.session_state:
        st.session_state.filename = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = None
    if 'cleaned_df' not in st.session_state:
        st.session_state.cleaned_df = None
    if 'validation_result' not in st.session_state:
        st.session_state.validation_result = None
    if 'aggregated_accounts' not in st.session_state:
        st.session_state.aggregated_accounts = None
    if 'processing_stats' not in st.session_state:
        st.session_state.processing_stats = None
    if 'processing_logs' not in st.session_state:
        st.session_state.processing_logs = []


def render_sidebar():
    """Render the navigation sidebar."""
    with st.sidebar:
        st.markdown(
            '<div class="sidebar-logo" aria-label="DataLens Cyber Cell Intelligence">'
            '<div class="sidebar-brand">'
            '<span class="brand-data">Data</span><span class="brand-lens">Lens</span>'
            '</div>'
            '<div class="brand-accent-line"><span></span><span></span><span></span><span></span></div>'
            '<div class="sidebar-subtitle">Cyber Cell Intelligence</div>'
            '</div>',
            unsafe_allow_html=True
        )
        st.markdown("---")
        
        # Page navigation
        pages = {
            'upload': '🔍 Aggregate by Account',
            'district_download': '📍 Victim-Suspect Mapping & Filter by State/District',
            'investment_fraud_top_accounts': 'Investment Fraud & Digital Arrest Top 30',
            'top_10_suspect': '🎯 Top 20 Suspect Accounts from Layer 1',
            'districtwise': '📊 Split Data by Column',
            'smart_district_split': '🗺️ Smart District Split',
            'ifsc_pincode_split': '🏦 Split Excel Districtwise by IFSC/PIN Code',
            'filter_by_entry_count': '🔢 Filter by Entry Count',
            'filter_by_unique_ack': '🏦 Filter by Entry Count (Unique ACK Only)',
            'non_gujarat_filter': '🗺️ Non-Gujarat Filter',
            'amount_matcher': '💰 Add Disputed Amount to Pending/Unattended ZIP File',
            'bank_ack_pivot': '🏦 Bank ACK Pivot',
            'ack_list_pivot': '📋 ACK List Pivot',
            'report_generator': '📊 Account & Hold Amount Report Generator',
            'automated_workflow': '🔄 Automated Workflow',
            'gujarat_account_formatter': 'Gujarat Unique Account Output',
            'kyc_notice_mailer': 'KYC Notice Mailer (NIC Mail)',
            'column_transfer': 'Add Columns by Match',
            'column_selector': '📋 Filter Excel File with the Columns You Need',
            'csv_fixer': 'CSV Fixer',
            'e_fir_report': 'E-FIR Report',
            'excel_merger': '📎 Merge Excel Files',
            'drop_call_finder': 'Drop Call Finder',
            'mo_finder': 'MO Finder',
            'call_notice_merge': '📞 Call Notice Data Mapping for Time Difference Between Call and Notice',
            'transaction_matcher': '🔄 Bring Disputed Amount to Unattended File (From Status-Wise Report)',
            'disputed_amount_matcher': '💰 Bring Disputed Amount to Any File with ACK No., Account Number & Transaction Amount',
            'money_transfer_dispute': '💸 Bring Disputed Amount to Money Transfer File',
            'ack_bank_consolidator': '📊 Acknowledgement Number & Bank-Wise Amount Consolidator',
            'bulk_mysql_import': '📊 Bulk MySQL Import',
            'mysql_database_viewer': '🗄️ MySQL Database Viewer',
            'ai_sql_assistant': '🤖 AI SQL Assistant',
            'distinct_account_pivot': '📊 Distinct Account Pivot',
            'view_database': '🗄️ View Database'
        }
        pages = dict(
            list(pages.items())[:6]
            + [('daily_report_district_split', 'Dailly Reports Split Suspect District Gujarat')]
            + list(pages.items())[6:]
        )

        clean_page_names = {
            'upload': 'Aggregate by Account',
            'district_download': 'Victim-Suspect Mapping',
            'investment_fraud_top_accounts': 'Investment Fraud & Digital Arrest Top 30',
            'top_10_suspect': 'Top 20 Suspect Accounts',
            'districtwise': 'Split Data by Column',
            'smart_district_split': 'Smart District Split',
            'ifsc_pincode_split': 'IFSC/PIN District Split',
            'filter_by_entry_count': 'Filter by Entry Count',
            'filter_by_unique_ack': 'Unique ACK Filter',
            'non_gujarat_filter': 'Non-Gujarat Filter',
            'amount_matcher': 'Disputed Amount Matcher',
            'bank_ack_pivot': 'Bank ACK Pivot',
            'ack_list_pivot': 'ACK List Pivot',
            'report_generator': 'Account & Hold Report',
            'automated_workflow': 'Automated Workflow',
            'gujarat_account_formatter': 'Gujarat Unique Account Output',
            'kyc_notice_mailer': 'KYC Notice Mailer (NIC Mail)',
            'column_transfer': 'Add Columns by Match',
            'column_selector': 'Column Selector',
            'csv_fixer': 'CSV Fixer',
            'e_fir_report': 'E-FIR Report',
            'excel_merger': 'Merge Excel Files',
            'drop_call_finder': 'Drop Call Finder',
            'mo_finder': 'MO Finder',
            'call_notice_merge': 'Call Notice Mapping',
            'transaction_matcher': 'Transaction Matcher',
            'disputed_amount_matcher': 'ACK Account Matcher',
            'money_transfer_dispute': 'Money Transfer Dispute',
            'ack_bank_consolidator': 'ACK Bank Consolidator',
            'bulk_mysql_import': 'Bulk MySQL Import',
            'mysql_database_viewer': 'MySQL Database Viewer',
            'ai_sql_assistant': 'AI SQL Assistant',
            'distinct_account_pivot': 'Distinct Account Pivot',
            'view_database': 'View Database'
        }

        clean_page_names['daily_report_district_split'] = 'Dailly Reports Split Suspect District Gujarat'

        top_page_keys = [
            'investment_fraud_top_accounts',
            'top_10_suspect',
            'automated_workflow',
            'gujarat_account_formatter',
            'kyc_notice_mailer',
            'column_transfer',
            'daily_report_district_split',
            'e_fir_report',
            'csv_fixer',
            'report_generator',
            'call_notice_merge',
            'drop_call_finder',
        ]

        def page_label(page_key):
            return clean_page_names.get(page_key, pages.get(page_key, page_key))

        def sidebar_section_title(title, detail=None):
            detail_html = f'<span>{detail}</span>' if detail else ''
            st.markdown(
                f'<div class="sidebar-section-title">{title}{detail_html}</div>',
                unsafe_allow_html=True,
            )

        sidebar_section_title("All Pages")
        ordered_page_keys = top_page_keys + [
            page_key
            for page_key in pages.keys()
            if page_key not in top_page_keys
        ]
        for page_key in ordered_page_keys:
            page_name = page_label(page_key)
            if st.button(page_name, key=f"nav_{page_key}", use_container_width=True):
                st.session_state.current_page = page_key
                st.rerun()
        
        st.markdown("---")
        
        # Session info
        st.caption("Session Info")
        if st.session_state.filename:
            st.caption(st.session_state.filename)
        
        st.markdown("---")
        
        # Data handling reminder
        st.caption("Security")
        st.caption("All data processed in-memory only.")
        st.caption("No data stored on servers.")
        
        # Reset button
        st.markdown("---")
        if st.button("Start Over", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()


def render_upload_page():
    """Render the file upload page with multiple file selection support."""
    services = get_services()
    upload_service = services['upload_service']
    
    # Render page header with info button
    render_page_header_with_info('upload')
    
    # Check if data already processed - show option to proceed or upload new
    if st.session_state.uploaded_df is not None:
        st.success(f"✅ Data already loaded: {len(st.session_state.uploaded_df)} rows from {st.session_state.filename}")
        
        if st.button("🔄 Upload New Files", use_container_width=True):
            st.session_state.uploaded_df = None
            st.session_state.filename = None
            st.session_state.column_mapping = None
            st.session_state.cleaned_df = None
            st.session_state.validation_result = None
            st.session_state.aggregated_accounts = None
            st.session_state.processing_stats = None
            st.rerun()
        
        # Show preview of current data
        st.markdown("---")
        st.subheader("📋 Current Data Preview")
        preview_df = upload_service.get_preview(st.session_state.uploaded_df, rows=10)
        st.dataframe(preview_df, use_container_width=True)
        return
    
    st.markdown("Upload **1 to 50 Excel/CSV files** using **Ctrl+Click** to select multiple files, then click **Process Files**.")
    
    # Multiple file uploader with Ctrl+Click support
    uploaded_files = st.file_uploader(
        "Choose Excel/CSV files (Ctrl+Click for multiple)",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        help="Supported formats: Excel (.xlsx, .xls) and CSV (.csv). Use Ctrl+Click to select multiple files."
    )
    
    # Show uploaded files list
    if uploaded_files:
        st.subheader(f"📁 {len(uploaded_files)} file(s) selected:")
        for f in uploaded_files:
            size_kb = f.size / 1024
            if size_kb > 1024:
                size_str = f"{size_kb/1024:.2f} MB"
            else:
                size_str = f"{size_kb:.2f} KB"
            st.write(f"• {f.name} — {size_str}")
        
        # Process Files button
        if st.button("🚀 Process Files", type="primary", use_container_width=True):
            if len(uploaded_files) > 50:
                st.warning("⚠️ Maximum 50 files allowed. Please remove some files.")
            else:
                all_data = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i, uploaded_file in enumerate(uploaded_files):
                    status_text.text(f"Reading file {i+1}/{len(uploaded_files)}: {uploaded_file.name}...")
                    progress_bar.progress((i + 1) / len(uploaded_files))
                    
                    try:
                        # Validate file
                        file_bytes = BytesIO(uploaded_file.getvalue())
                        validation_result = upload_service.validate_file(file_bytes, uploaded_file.name)
                        
                        if not validation_result.is_valid:
                            st.warning(f"⚠️ Skipping {uploaded_file.name}: {validation_result.error_message}")
                            continue
                        
                        # Read file
                        file_bytes.seek(0)
                        df = upload_service.read_file(file_bytes, uploaded_file.name)
                        
                        st.success(f"✅ {uploaded_file.name}: {len(df)} rows loaded")
                        all_data.append((uploaded_file.name, df))
                        
                    except Exception as e:
                        st.error(f"❌ Error reading {uploaded_file.name}: {str(e)}")
                
                progress_bar.progress(100)
                status_text.text("✅ All files processed!")
                
                if all_data:
                    # Combine all dataframes
                    combined_df = pd.concat([df for _, df in all_data], ignore_index=True)
                    
                    # Store in session state immediately
                    st.session_state.uploaded_df = combined_df
                    st.session_state.filename = f"{len(all_data)}_files_combined"
                    
                    # Rerun to show the proceed option
                    st.rerun()
                else:
                    st.error("❌ No valid data found in uploaded files.")
    else:
        st.info("📤 Select files using Ctrl+Click for multiple selection, then click **Process Files**.")


def render_mapping_page():
    """Render the column mapping page."""
    services = get_services()
    column_detector = services['column_detector']
    
    if st.session_state.uploaded_df is None:
        st.warning("Please upload a file first.")
        return
    
    # Render page header with info button
    render_page_header_with_info('mapping')
    
    df = st.session_state.uploaded_df
    headers = list(df.columns)
    
    # Auto-detect columns
    auto_mapping = column_detector.detect_columns(headers)
    
    # Show confidence scores
    if auto_mapping.confidence_scores:
        st.subheader("🎯 Auto-Detection Confidence")
        conf_cols = st.columns(4)
        for i, (col_type, score) in enumerate(auto_mapping.confidence_scores.items()):
            with conf_cols[i % 4]:
                color = "green" if score >= 0.9 else "orange" if score >= 0.8 else "red"
                st.markdown(f"**{col_type.replace('_', ' ').title()}**: :{color}[{score:.0%}]")
    
    # Show ambiguous mappings warning
    if auto_mapping.ambiguous_mappings:
        st.warning("⚠️ Some columns have ambiguous mappings. Please verify the selections below.")
    
    st.markdown("---")
    st.subheader("📝 Column Assignments")
    
    # Create mapping form
    col1, col2 = st.columns(2)
    
    # Add "None" option to headers
    header_options = ["-- Not Mapped --"] + headers
    
    def get_default_index(mapped_value):
        if mapped_value and mapped_value in headers:
            return headers.index(mapped_value) + 1
        return 0
    
    with col1:
        bank_account_col = st.selectbox(
            "Bank Account Number",
            options=header_options,
            index=get_default_index(auto_mapping.bank_account_number),
            help="Column containing fraudster bank account numbers"
        )
        
        amount_col = st.selectbox(
            "Amount",
            options=header_options,
            index=get_default_index(auto_mapping.amount),
            help="Column containing transaction amounts"
        )
        
        ack_col = st.selectbox(
            "Acknowledgement Number",
            options=header_options,
            index=get_default_index(auto_mapping.acknowledgement_number),
            help="Column containing acknowledgement/reference numbers"
        )
        
        ifsc_col = st.selectbox(
            "IFSC Code",
            options=header_options,
            index=get_default_index(auto_mapping.ifsc_code),
            help="Column containing bank IFSC codes"
        )
    
    with col2:
        bank_name_col = st.selectbox(
            "Bank Name",
            options=header_options,
            index=get_default_index(auto_mapping.bank_name),
            help="Column containing bank names"
        )
        
        address_col = st.selectbox(
            "Address",
            options=header_options,
            index=get_default_index(auto_mapping.address),
            help="Column containing beneficiary addresses"
        )
        
        disputed_col = st.selectbox(
            "Disputed Amount",
            options=header_options,
            index=get_default_index(auto_mapping.disputed_amount),
            help="Column containing disputed/chargeback amounts"
        )
        
        serial_col = st.selectbox(
            "Serial Number",
            options=header_options,
            index=get_default_index(auto_mapping.serial_number),
            help="Column containing serial/row numbers"
        )
        
        district_col = st.selectbox(
            "District",
            options=header_options,
            index=get_default_index(auto_mapping.district),
            help="Column containing district names"
        )
        
        state_col = st.selectbox(
            "State",
            options=header_options,
            index=get_default_index(auto_mapping.state),
            help="Column containing state names"
        )


    # Validation
    st.markdown("---")
    
    st.success("✅ Map the columns that are available in your file")
    
    # Proceed button
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("Proceed to Processing", type="primary", use_container_width=True):
            # Create final mapping
            final_mapping = ColumnMapping(
                serial_number=serial_col if serial_col != "-- Not Mapped --" else None,
                acknowledgement_number=ack_col if ack_col != "-- Not Mapped --" else None,
                bank_account_number=bank_account_col if bank_account_col != "-- Not Mapped --" else None,
                ifsc_code=ifsc_col if ifsc_col != "-- Not Mapped --" else None,
                address=address_col if address_col != "-- Not Mapped --" else None,
                amount=amount_col if amount_col != "-- Not Mapped --" else None,
                disputed_amount=disputed_col if disputed_col != "-- Not Mapped --" else None,
                bank_name=bank_name_col if bank_name_col != "-- Not Mapped --" else None,
                district=district_col if district_col != "-- Not Mapped --" else None,
                state=state_col if state_col != "-- Not Mapped --" else None
            )
            
            st.session_state.column_mapping = final_mapping
            st.rerun()


def render_processing_page():
    """Render the data processing page with progress tracking."""
    services = get_services()
    data_processor = services['data_processor']
    validation_engine = services['validation_engine']
    aggregation_engine = services['aggregation_engine']
    dashboard = services['dashboard']
    
    if st.session_state.uploaded_df is None or st.session_state.column_mapping is None:
        st.warning("Please complete the previous steps first.")
        return
    
    # Render page header with info button
    render_page_header_with_info('processing')
    
    df = st.session_state.uploaded_df
    mapping = st.session_state.column_mapping
    
    # Check if already processed
    if st.session_state.aggregated_accounts is not None:
        st.success("✅ Data has already been processed!")
        
        if st.button("Reprocess Data", use_container_width=True):
            st.session_state.aggregated_accounts = None
            st.session_state.processing_stats = None
            st.session_state.cleaned_df = None
            st.session_state.validation_result = None
            st.session_state.processing_logs = []
            st.rerun()
        return
    
    # Processing controls
    if st.button("🚀 Start Processing", type="primary", use_container_width=True):
        logs = []
        
        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.container()
        
        # Metrics placeholders
        metrics_cols = st.columns(4)
        rows_metric = metrics_cols[0].empty()
        accounts_metric = metrics_cols[1].empty()
        errors_metric = metrics_cols[2].empty()
        amount_metric = metrics_cols[3].empty()
        
        total_rows = len(df)
        
        try:
            # Step 1: Data Cleaning (20%)
            status_text.text("Step 1/4: Cleaning data...")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Starting data cleaning...")
            
            cleaned_df = data_processor.clean_dataframe(df, mapping)
            rows_after_cleaning = len(cleaned_df)
            
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Removed {total_rows - rows_after_cleaning} empty rows")
            progress_bar.progress(20)
            rows_metric.metric("Rows Processed", rows_after_cleaning)
            
            # Step 2: Validation (40%)
            status_text.text("Step 2/4: Validating data...")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Starting data validation...")
            
            validation_result = validation_engine.validate_dataframe(cleaned_df, mapping)
            
            if not validation_result.is_valid:
                for error in validation_result.critical_errors:
                    logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] CRITICAL: {error}")
                st.error("❌ Critical validation errors found. Cannot proceed.")
                for error in validation_result.critical_errors:
                    st.error(error)
                return
            
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Found {len(validation_result.warnings)} warnings")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Flagged {len(validation_result.flagged_rows)} rows")
            progress_bar.progress(40)
            errors_metric.metric("Errors Found", len(validation_result.flagged_rows))


            # Step 3: Aggregation (70%)
            status_text.text("Step 3/4: Aggregating by account...")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Starting aggregation by account number...")
            
            aggregated = aggregation_engine.aggregate_by_account(cleaned_df, mapping)
            sorted_accounts = aggregation_engine.sort_results(aggregated)
            
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Found {len(sorted_accounts)} unique accounts")
            progress_bar.progress(70)
            accounts_metric.metric("Unique Accounts", len(sorted_accounts))
            
            # Step 4: Calculate Statistics (100%)
            status_text.text("Step 4/4: Calculating statistics...")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Calculating summary statistics...")
            
            stats = dashboard.calculate_statistics(
                accounts=sorted_accounts,
                total_input_rows=total_rows,
                input_filename=st.session_state.filename or "",
                rows_with_errors=len(validation_result.flagged_rows)
            )
            
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Total fraud amount: ₹{stats.total_fraud_amount:,.2f}")
            progress_bar.progress(100)
            amount_metric.metric("Total Amount", f"₹{stats.total_fraud_amount:,.0f}")
            
            # Store results
            st.session_state.cleaned_df = cleaned_df
            st.session_state.validation_result = validation_result
            st.session_state.aggregated_accounts = sorted_accounts
            st.session_state.processing_stats = stats
            st.session_state.processing_logs = logs
            
            status_text.text("✅ Processing complete!")
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] Processing completed successfully!")
            
            st.success("🎉 Processing completed successfully!")
            
            # Show logs
            with log_container:
                st.subheader("📜 Processing Log")
                log_text = "\n".join(logs)
                st.text_area("Logs", value=log_text, height=200, disabled=True)
                
        except Exception as e:
            logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] ERROR: {str(e)}")
            st.error(f"❌ Processing failed: {str(e)}")
            
            with log_container:
                st.subheader("📜 Processing Log")
                log_text = "\n".join(logs)
                st.text_area("Logs", value=log_text, height=200, disabled=True)


def render_results_page():
    """Render the results dashboard with statistics, downloads, and filters."""
    services = get_services()
    dashboard = services['dashboard']
    report_generator = services['report_generator']
    
    if st.session_state.aggregated_accounts is None:
        st.warning("Please process data first.")
        return
    
    # Render page header with info button
    render_page_header_with_info('results')
    
    accounts = st.session_state.aggregated_accounts
    stats = st.session_state.processing_stats
    validation_result = st.session_state.validation_result

    def safe_text(value):
        if value is None:
            return ""
        try:
            if value != value:
                return ""
        except Exception:
            pass
        text = str(value).strip()
        if text.lower() in {"nan", "none", "<na>", "nat"}:
            return ""
        return text

    def truncate_text(value, max_length):
        text = safe_text(value)
        return text[:max_length] + "..." if len(text) > max_length else text
    
    # Summary Statistics
    st.subheader("📈 Summary Statistics")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Input Rows", stats.total_input_rows)
    with col2:
        st.metric("Unique Accounts", stats.unique_accounts)
    with col3:
        st.metric("Total Fraud Amount", f"₹{stats.total_fraud_amount:,.2f}")
    with col4:
        st.metric("Avg Amount/Account", f"₹{stats.average_amount_per_account:,.2f}")
    
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("Rows Processed", stats.rows_processed)
    with col6:
        st.metric("Rows with Errors", stats.rows_with_errors)
    with col7:
        st.metric("Total Disputed", f"₹{stats.total_disputed_amount:,.2f}")
    with col8:
        if stats.unique_accounts > 0:
            avg_txn = stats.total_input_rows / stats.unique_accounts
            st.metric("Avg Txn/Account", f"{avg_txn:.1f}")
        else:
            st.metric("Avg Txn/Account", "0")
    
    st.markdown("---")
    
    # Download Section - Lazy generation (no eager computation)
    st.subheader("📥 Download Reports")
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    quality_metrics = validation_result.quality_report if validation_result else None
    errors = validation_result.warnings if validation_result else []
    
    # Row 1: Download buttons
    dl_col1, dl_col2, dl_col3, dl_col4 = st.columns(4)
    
    with dl_col1:
        if st.button("📊 Prepare Excel", use_container_width=True, key="prep_excel"):
            with st.spinner("Generating Excel..."):
                st.session_state.excel_ready = report_generator.generate_excel_bytes(accounts)
        
        if 'excel_ready' in st.session_state and st.session_state.excel_ready:
            st.download_button(
                label="⬇️ Download Excel",
                data=st.session_state.excel_ready,
                file_name=f"fraud_analysis_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    with dl_col2:
        if st.button("📄 Prepare CSV", use_container_width=True, key="prep_csv"):
            with st.spinner("Generating CSV..."):
                st.session_state.csv_ready = report_generator.generate_csv_bytes(accounts)
        
        if 'csv_ready' in st.session_state and st.session_state.csv_ready:
            st.download_button(
                label="⬇️ Download CSV",
                data=st.session_state.csv_ready,
                file_name=f"fraud_analysis_{timestamp}.csv",
                mime="text/csv",
                use_container_width=True
            )
    
    with dl_col3:
        if st.button("📑 Prepare PDF", use_container_width=True, key="prep_pdf"):
            with st.spinner("Generating PDF..."):
                st.session_state.pdf_ready = report_generator.generate_pdf_bytes(accounts, stats, quality_metrics)
        
        if 'pdf_ready' in st.session_state and st.session_state.pdf_ready:
            st.download_button(
                label="⬇️ Download PDF",
                data=st.session_state.pdf_ready,
                file_name=f"fraud_analysis_{timestamp}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
    
    with dl_col4:
        if st.button("📋 Prepare Audit Log", use_container_width=True, key="prep_audit"):
            with st.spinner("Generating Audit Log..."):
                audit_log = report_generator.generate_audit_log(
                    input_filename=st.session_state.filename or "",
                    rows_processed=stats.rows_processed,
                    errors_encountered=errors,
                    timestamp=stats.processing_timestamp
                )
                st.session_state.audit_ready = audit_log.encode('utf-8')
        
        if 'audit_ready' in st.session_state and st.session_state.audit_ready:
            st.download_button(
                label="⬇️ Download Audit Log",
                data=st.session_state.audit_ready,
                file_name=f"audit_log_{timestamp}.txt",
                mime="text/plain",
                use_container_width=True
            )
    
    # Row 2: SAVE TO DATABASE - Prominent button
    st.markdown("")
    st.markdown("##### 🗄️ Save to MySQL Database")
    
    db_save_col1, db_save_col2, db_save_col3 = st.columns([2, 1, 1])
    
    with db_save_col1:
        save_dataset_name = st.text_input("Dataset Name", placeholder="e.g., January 2025 Fraud Data", key="save_ds_name_input")
    
    with db_save_col2:
        st.markdown("<br>", unsafe_allow_html=True)
        save_to_db_btn = st.button("💾 SAVE TO DATABASE", use_container_width=True, key="save_to_db_main", type="primary")
    
    with db_save_col3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⚙️ DB Settings", use_container_width=True, key="db_settings_btn"):
            st.session_state.show_db_settings = True
    
    # Database settings expander
    if st.session_state.get('show_db_settings', False):
        with st.expander("Database Connection Settings", expanded=True):
            db_set_col1, db_set_col2 = st.columns(2)
            with db_set_col1:
                st.session_state.db_host = st.text_input("Host", value=st.session_state.get('db_host', 'localhost'), key="db_host_set")
                st.session_state.db_user = st.text_input("Username", value=st.session_state.get('db_user', 'root'), key="db_user_set")
            with db_set_col2:
                st.session_state.db_port = st.number_input("Port", value=st.session_state.get('db_port', 3306), key="db_port_set")
                st.session_state.db_password = st.text_input("Password", value=st.session_state.get('db_password', 'Cyber2026'), type="password", key="db_pass_set")
            
            if st.button("🔌 Test Connection", key="test_conn_btn"):
                db_service = DatabaseService(
                    host=st.session_state.get('db_host', 'localhost'),
                    port=st.session_state.get('db_port', 3306),
                    user=st.session_state.get('db_user', 'root'),
                    password=st.session_state.get('db_password', 'Cyber2026')
                )
                success, msg = db_service.test_connection()
                if success:
                    st.success(f"✅ {msg}")
                else:
                    st.error(f"❌ {msg}")
    
    # Handle Save to Database
    if save_to_db_btn:
        if not save_dataset_name:
            st.error("⚠️ Please enter a dataset name!")
        else:
            with st.spinner("💾 Saving to MySQL database..."):
                db_service = DatabaseService(
                    host=st.session_state.get('db_host', 'localhost'),
                    port=st.session_state.get('db_port', 3306),
                    user=st.session_state.get('db_user', 'root'),
                    password=st.session_state.get('db_password', 'Cyber2026')
                )
                
                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def update_progress(current, total):
                    progress = int((current / total) * 100)
                    progress_bar.progress(progress)
                    status_text.text(f"Saving... {current:,} / {total:,} records ({progress}%)")
                
                dataset_id, error_msg = db_service.save_dataset(
                    name=save_dataset_name,
                    description=f"Saved from DataLens for Cyber Cell - {len(accounts)} accounts",
                    accounts=accounts,
                    source_filename=st.session_state.get('filename', ''),
                    progress_callback=update_progress
                )
                db_service.disconnect()
                
                progress_bar.empty()
                status_text.empty()
                
                if dataset_id:
                    st.success(f"""
                    ✅ **Data Saved Successfully!**
                    - Dataset ID: {dataset_id}
                    - Records: {len(accounts):,}
                    - Database: gujarat_cyber_police
                    - You can view this in MySQL Workbench
                    """)
                else:
                    st.error(f"❌ Save failed: {error_msg}")


# Main execution
init_session_state()
render_sidebar()

# Route to the appropriate page based on session state
current_page = st.session_state.current_page

if current_page == 'upload':
    render_upload_page()
elif current_page == 'district_download':
    render_district_download_page()
elif current_page == 'investment_fraud_top_accounts':
    render_investment_fraud_top_accounts_page()
elif current_page == 'top_10_suspect':
    render_top_10_suspect_accounts_page()
elif current_page == 'districtwise':
    render_districtwise_page()
elif current_page == 'smart_district_split':
    render_smart_district_split_page()
elif current_page == 'ifsc_pincode_split':
    render_ifsc_pincode_district_split_page()
elif current_page == 'daily_report_district_split':
    render_daily_report_district_split_page()
elif current_page == 'e_fir_report':
    render_efir_report_page()
elif current_page == 'filter_by_entry_count':
    render_filter_by_entry_count_page()
elif current_page == 'filter_by_unique_ack':
    render_filter_by_unique_ack_page()
elif current_page == 'non_gujarat_filter':
    render_non_gujarat_filter_page()
elif current_page == 'amount_matcher':
    render_amount_matcher_page()
elif current_page == 'bank_ack_pivot':
    render_bank_ack_pivot_page()
elif current_page == 'ack_list_pivot':
    render_ack_list_pivot_page()
elif current_page == 'report_generator':
    from src.report_service import render_report_generator_page
    render_report_generator_page()
elif current_page == 'automated_workflow':
    render_automated_workflow_page()
elif current_page == 'gujarat_account_formatter':
    render_gujarat_account_formatter_page()
elif current_page == 'kyc_notice_mailer':
    render_kyc_notice_mailer_page()
elif current_page == 'column_transfer':
    render_column_transfer_page()
elif current_page == 'column_selector':
    render_column_selector_page()
elif current_page == 'csv_fixer':
    render_csv_fixer_page()
elif current_page == 'excel_merger':
    render_excel_merger_page()
elif current_page == 'drop_call_finder':
    render_drop_call_finder_page()
elif current_page == 'mo_finder':
    render_mo_finder_page()
elif current_page == 'call_notice_merge':
    render_call_notice_merge_page()
elif current_page == 'transaction_matcher':
    render_transaction_matcher_page()
elif current_page == 'disputed_amount_matcher':
    render_disputed_amount_matcher_page()
elif current_page == 'money_transfer_dispute':
    render_money_transfer_dispute_page()
elif current_page == 'ack_bank_consolidator':
    render_ack_bank_consolidator_page()
elif current_page == 'bulk_mysql_import':
    render_bulk_mysql_import_page()
elif current_page == 'mysql_database_viewer':
    render_mysql_database_viewer_page()
elif current_page == 'ai_sql_assistant':
    render_ai_sql_assistant_page()
elif current_page == 'distinct_account_pivot':
    render_distinct_account_pivot_page()
elif current_page == 'view_database':
    from src.database_service import DatabaseService
    st.header("🗄️ View Database")
    st.info("Database viewer coming soon...")
else:
    st.error(f"Unknown page: {current_page}")
    st.info("Redirecting to home page...")
    st.session_state.current_page = 'upload'
    st.rerun()
