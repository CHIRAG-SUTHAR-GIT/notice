"""
UI Styling and Theme Configuration for DataLens for Cyber Cell
Hyper-Premium Midnight Theme with Ultra Smooth Animations
"""

from html import escape

import streamlit as st

# Color Scheme - Cyber intelligence report theme
COLORS = {
    # Primary - matte dark interface
    'deep_ocean': '#080B10',
    'ocean_blue': '#10141B',
    'water_blue': '#2F80FF',
    'sky_blue': '#6F8BAE',
    'light_aqua': '#F04348',
    'foam_white': '#F8F3EA',
    
    # Accent - report ranking colors
    'coral': '#F04348',
    'teal': '#21C16B',
    'mint': '#FF9F0A',
    'seafoam': '#2F80FF',
    
    # Background
    'background': '#070A0F',
    'background_card': '#11161E',
    'text_primary': '#F8F3EA',
    'text_secondary': '#91A0B8',
    'text_muted': '#6F7D92',
    
    # Status colors
    'success': '#21C16B',
    'warning': '#FF9F0A',
    'error': '#F04348',
    'info': '#2F80FF',
    
    # Panel effect
    'glass_bg': '#11161E',
    'glass_border': '#242C39',
    
    # Additional
    'border': '#242C39',
    'hover': '#181E28',
    'shadow': 'rgba(0, 0, 0, 0.35)'
}

PAGE_INFO = {
    'upload': {
        'title': '📤 Upload Transaction Files',
        'description': 'Upload Excel or CSV files containing fraud transaction data.',
        'details': '''
        **What this page does:**
        - Accepts Excel (.xlsx, .xls) and CSV files
        - Supports uploading 1-50 files at once
        - Automatically combines multiple files
        '''
    },
    'mapping': {
        'title': '🔗 Column Mapping',
        'description': 'Map your file columns to required fields.',
        'details': '''
        **What this page does:**
        - Auto-detects column mappings
        - Allows manual column selection
        '''
    },
    'processing': {
        'title': '⚙️ Data Processing',
        'description': 'Process and validate uploaded data.',
        'details': '''
        **What this page does:**
        - Data cleaning and validation
        - Aggregation by account number
        '''
    },
    'results': {
        'title': '📊 Results Dashboard',
        'description': 'View processed results with statistics.',
        'details': '''
        **Features:**
        - Summary statistics
        - Search and filter
        - Download reports
        '''
    },
    'district_download': {
        'title': '📍 District Data Download',
        'description': 'Download data filtered by district.',
        'details': '''
        **What this page does:**
        - Filter by district
        - Download district-specific data
        '''
    },
    'investment_fraud_top_accounts': {
        'title': 'Investment Fraud & Digital Arrest Top 30',
        'description': 'Rank Layer 1 bank accounts connected to investment-fraud and digital-arrest complaints.',
        'details': '''
        **What this page does:**
        - Detects investment-fraud and digital-arrest complaints from the Additional Information report
        - Matches acknowledgement numbers to the Layerwise report
        - Keeps numeric Layer 1 bank accounts and combines repeated account numbers
        - Ranks the top 30 by total disputed amount
        - Creates a formatted Excel report for download
        '''
    },
    'e_fir_report': {
        'title': 'E-FIR Report',
        'description': 'Match Crime Report and Zero E-FIR records by acknowledgement number.',
        'details': '''
        **What this page does:**
        - Filters the E-FIR Details sheet by the selected E-FIR date
        - Matches acknowledgement numbers across both reports
        - Uses Total Reported Amount from the E-FIR Details sheet
        - Keeps one Crime Report record for each matched acknowledgement number
        - Sorts by Date of e-Zero FIR time, with the oldest record first
        - Creates a formatted E-FIR Excel report for download
        '''
    },
    'districtwise': {
        'title': '📊 Split Data by Column',
        'description': 'Split your data into separate files based on column values.',
        'details': '''
        **What this page does:**
        - Split by any column
        - Create separate files or sheets
        '''
    },
    'smart_district_split': {
        'title': '🗺️ Smart District Split',
        'description': 'Intelligently split data by Gujarat districts.',
        'details': '''
        **What this page does:**
        - Maps talukas to districts
        - Handles spelling variations
        '''
    },
    'ifsc_pincode_split': {
        'title': '🏦 IFSC/PIN District Split',
        'description': '100% accurate district identification using IFSC and PIN codes.',
        'details': '''
        **What this page does:**
        - IFSC code lookup (real-time API)
        - PIN code mapping
        - 6-layer identification system
        '''
    },
    'filter_by_entry_count': {
        'title': '🔢 Filter by Entry Count',
        'description': 'Filter data to show only entries that appear a minimum number of times.',
        'details': '''
        **What this page does:**
        - Groups data by selected column
        - Filters by minimum entry count
        '''
    },
    'filter_by_unique_ack': {
        'title': '🏦 Filter Banks by Unique ACK',
        'description': 'Filter banks based on unique ACK count.',
        'details': '''
        **What this page does:**
        - Count unique ACKs per bank
        - Filter banks with minimum unique ACKs
        - Output all records including duplicates
        '''
    },
    'non_gujarat_filter': {
        'title': '🗺️ Non-Gujarat Filter',
        'description': 'Filter out Gujarat state data.',
        'details': '''
        **What this page does:**
        - Removes Gujarat entries
        - Keeps only Non-Gujarat states
        '''
    },
    'amount_matcher': {
        'title': '💰 Amount Matcher',
        'description': 'Match transactions by amount.',
        'details': '''
        **What this page does:**
        - Match records by amount
        - Find duplicate amounts
        '''
    },
    'bank_ack_pivot': {
        'title': '🏦 Bank ACK Pivot',
        'description': 'Create pivot table of banks and ACK numbers.',
        'details': '''
        **What this page does:**
        - Pivot by bank and ACK
        - Summary statistics
        '''
    },
    'ack_list_pivot': {
        'title': '📋 ACK List Pivot',
        'description': 'Create pivot table of ACK numbers.',
        'details': '''
        **What this page does:**
        - Pivot by ACK number
        - Detailed breakdown
        '''
    },
    'automated_workflow': {
        'title': '🔄 Automated Workflow',
        'description': 'Automated processing workflow.',
        'details': '''
        **What this page does:**
        - End-to-end automation
        - Batch processing
        '''
    },
    'column_selector': {
        'title': '📋 Column Selector',
        'description': 'Select and reorder columns.',
        'details': '''
        **What this page does:**
        - Choose columns to keep
        - Reorder columns
        '''
    },
    'excel_merger': {
        'title': '📎 Merge Excel Files',
        'description': 'Merge multiple Excel files into one.',
        'details': '''
        **What this page does:**
        - Combine multiple files
        - Preserve all data
        '''
    },
    'call_notice_merge': {
        'title': '📞 Call Notice Data Merge',
        'description': 'Merge call notice data.',
        'details': '''
        **What this page does:**
        - Merge call and notice data
        - Match by key fields
        '''
    },
    'transaction_matcher': {
        'title': '🔄 Transaction Matcher',
        'description': 'Match transactions by Transaction ID.',
        'details': '''
        **What this page does:**
        - Match by Transaction ID
        - Add account and date info
        '''
    },
    'disputed_amount_matcher': {
        'title': '💰 Disputed Amount Matcher',
        'description': 'Match disputed amounts.',
        'details': '''
        **What this page does:**
        - Match disputed amounts
        - Multi-parameter matching
        '''
    },
    'money_transfer_dispute': {
        'title': '💸 Money Transfer Dispute Matcher',
        'description': 'Add Disputed Amount to Money Transfer files.',
        'details': '''
        **What this page does:**
        - Match by Ack No, Account No, and Amount
        - Uses backward matching for account numbers
        '''
    },
    'ack_bank_consolidator': {
        'title': '📊 ACK + Bank Consolidator',
        'description': 'Consolidate rows by ACK Number + Bank Name.',
        'details': '''
        **What this page does:**
        - Groups rows by ACK Number + Bank Name
        - Sums transaction amounts per group
        - Reduces multiple entries to one per bank per ACK
        '''
    },
    'bulk_mysql_import': {
        'title': '📊 Bulk MySQL Import',
        'description': 'Import all Excel files from a folder into MySQL database.',
        'details': '''
        **What this page does:**
        - Scan folder for Excel files
        - Automatically create MySQL tables
        - Import all data in batches
        - Handle large files efficiently
        
        **Features:**
        - Batch import (1000 rows at a time)
        - Auto table name sanitization
        - Data type detection
        - Progress tracking
        - Error handling
        '''
    },
    'mysql_database_viewer': {
        'title': '🗄️ MySQL Database Viewer',
        'description': 'Browse, search, filter, and download data from MySQL tables.',
        'details': '''
        **What this page does:**
        - Connect to MySQL database
        - View all tables with row counts
        - Browse table data with pagination
        - Search and filter records
        - Download data in Excel/CSV
        
        **Features:**
        - Table browser with statistics
        - Search across all columns
        - Sort and filter data
        - Column selection
        - Download current view or entire table
        - Data statistics and insights
        '''
    },
    'ai_sql_assistant': {
        'title': '🤖 AI SQL Assistant',
        'description': 'Ask questions in natural language and get SQL queries automatically.',
        'details': '''
        **What this page does:**
        - Convert natural language to SQL
        - Auto-generate queries from questions
        - Execute queries and show results
        - Download query results
        
        **Features:**
        - Natural language processing
        - Supports counting, aggregation, searching
        - Query editing and history
        - Quick question buttons
        - Excel/CSV export
        - No SQL knowledge required
        '''
    },
    'report_generator': {
        'title': '📊 Account & Hold Amount Report Generator',
        'description': 'Generate professional reports with Account, Hold Amount, and Unattended data',
        'details': '''
        **Report Generator Features:**
        - Three comprehensive sheets in one Excel file
        - PUT ON HOLD: Bank-wise hold amounts with grand total
        - ACCOUNT: Bank-wise account count (distinct ACK numbers)
        - Complaint Un Attended: Bank/Wallet/Merchant-wise unattended complaints
        
        **Professional Formatting:**
        - Sky blue title rows
        - Light yellow headers
        - Table format with borders
        - Grand total rows in bold with sky blue background
        
        **Data Processing:**
        - Automatic bank name filtering
        - Removal of "other/others/othar" entries
        - Zero value filtering
        - Distinct count calculations
        - Bank-wise aggregations
        '''
    }
}

def apply_custom_css():
    """Apply custom CSS styling to the Streamlit app - Hyper-Premium Midnight Theme"""
    st.markdown(f"""
    <style>
    /* Import Modern Fonts - Ultra Premium Typography */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@200;300;400;500;600;700;800&family=JetBrains+Mono:wght@400;700&display=swap');
    
    /* Global Styles */
    * {{
        font-family: 'Outfit', sans-serif;
        scroll-behavior: smooth !important;
    }}
    
    /* Remove White Gap at Top */
    header[data-testid="stHeader"] {{
        background: transparent !important;
        height: 0px !important;
        padding: 0 !important;
    }}
    
    /* Ensure no white spots appear */
    .stApp > header {{
        background-color: transparent !important;
    }}
    
    [data-testid="stAppViewContainer"] {{
        background-color: transparent !important;
    }}
    
    /* Main Background - Liquid Dark Gradient */
    .stApp {{
        background: radial-gradient(circle at 50% 10%, {COLORS['ocean_blue']} 0%, {COLORS['deep_ocean']} 60%, {COLORS['background']} 100%) !important;
        background-attachment: fixed !important;
        color: {COLORS['text_primary']};
    }}
    
    /* Main Container Glass Scrolling Effect - Like sliding on glass */
    [data-testid="stAppViewBlockContainer"] {{
        background: rgba(10, 10, 14, 0.5) !important;
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
        border: 1px solid rgba(0, 245, 255, 0.15) !important;
        border-radius: 24px !important;
        padding: 3rem !important;
        margin-top: 3rem !important;
        margin-bottom: 3rem !important;
        box-shadow: 0 25px 50px rgba(0,0,0,0.6), inset 0 2px 0 rgba(255,255,255,0.1) !important;
        max-width: 95% !important;
        transition: all 0.4s ease;
    }}
    
    /* Smooth Scrollbar for entire app */
    ::-webkit-scrollbar {{
        width: 10px;
        height: 10px;
    }}
    ::-webkit-scrollbar-track {{
        background: {COLORS['background']};
        border-radius: 8px;
    }}
    ::-webkit-scrollbar-thumb {{
        background: linear-gradient(180deg, {COLORS['sky_blue']} 0%, {COLORS['water_blue']} 100%);
        border-radius: 8px;
        border: 2px solid {COLORS['background']};
        transition: background 0.5s cubic-bezier(0.2, 0.8, 0.2, 1);
    }}
    ::-webkit-scrollbar-thumb:hover {{
        background: linear-gradient(180deg, {COLORS['teal']} 0%, {COLORS['seafoam']} 100%);
        box-shadow: 0 0 15px {COLORS['teal']};
    }}
    
    /* Liquid Glow Animation */
    @keyframes liquidGlow {{
        0% {{ box-shadow: 0 0 10px rgba(0, 245, 255, 0.2), inset 0 0 15px rgba(0, 245, 255, 0.1); }}
        50% {{ box-shadow: 0 0 25px rgba(0, 245, 255, 0.5), inset 0 0 20px rgba(0, 245, 255, 0.3); }}
        100% {{ box-shadow: 0 0 10px rgba(0, 245, 255, 0.2), inset 0 0 15px rgba(0, 245, 255, 0.1); }}
    }}
    
    @keyframes floatingDrops {{
        0% {{ transform: translateY(0px) scale(1); opacity: 0.1; }}
        50% {{ transform: translateY(-30px) scale(1.1); opacity: 0.3; }}
        100% {{ transform: translateY(0px) scale(1); opacity: 0.1; }}
    }}
    
    /* Sidebar - Glass Container */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, rgba(4, 4, 6, 0.95) 0%, rgba(10, 10, 14, 0.98) 100%) !important;
        backdrop-filter: blur(30px) saturate(200%) !important;
        -webkit-backdrop-filter: blur(30px) saturate(200%) !important;
        border-right: 2px solid rgba(0, 245, 255, 0.5) !important;
        box-shadow: 10px 0 40px rgba(0, 0, 0, 0.8), 2px 0 15px rgba(0, 245, 255, 0.2) !important;
        position: relative;
        overflow: hidden;
    }}
    
    /* Adding the sliding glass look to sidebar header */
    [data-testid="stSidebarNav"] {{
        background: transparent !important;
    }}
    
    [data-testid="stSidebar"]::before {{
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: 
            radial-gradient(circle at 10% 20%, {COLORS['seafoam']} 0%, transparent 8%),
            radial-gradient(circle at 80% 60%, {COLORS['sky_blue']} 0%, transparent 5%),
            radial-gradient(circle at 40% 90%, {COLORS['teal']} 0%, transparent 6%);
        filter: blur(20px);
        opacity: 0.2;
        animation: floatingDrops 12s infinite alternate ease-in-out;
        pointer-events: none;
        z-index: 0;
    }}
    
    [data-testid="stSidebar"] > div {{ position: relative; z-index: 1; }}
    
    /* Sidebar Navigation Buttons - Holographic Water Effect */
    [data-testid="stSidebar"] .stButton button {{
        background: padding-box linear-gradient(rgba(22, 22, 26, 0.4), rgba(22, 22, 26, 0.4)),
                    border-box linear-gradient(135deg, {COLORS['glass_border']}, transparent, {COLORS['glass_border']});
        color: {COLORS['light_aqua']};
        border: 1px solid transparent;
        border-radius: 12px;
        font-weight: 500;
        font-size: 0.95rem;
        padding: 14px 22px;
        transition: all 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        backdrop-filter: blur(10px);
        box-shadow: inset 0 0 10px rgba(0, 245, 255, 0.05);
        position: relative;
        overflow: hidden;
        text-transform: uppercase;
        letter-spacing: 1px;
    }}
    
    /* Liquid Hover Ripple Effect */
    [data-testid="stSidebar"] .stButton button::before {{
        content: '';
        position: absolute;
        bottom: 0; left: 0; right: 0;
        height: 0%;
        background: linear-gradient(180deg, rgba(0, 245, 255, 0), rgba(0, 245, 255, 0.4));
        transition: height 0.4s ease-out;
        z-index: 0;
        border-radius: 10px;
    }}
    
    [data-testid="stSidebar"] .stButton button:hover::before {{ height: 100%; }}
    
    [data-testid="stSidebar"] .stButton button:hover {{
        color: {COLORS['deep_ocean']};
        background: linear-gradient(135deg, {COLORS['sky_blue']} 0%, {COLORS['teal']} 100%);
        border-color: {COLORS['seafoam']};
        transform: translateY(-3px) scale(1.02);
        box-shadow: 
            0 10px 25px rgba(0, 245, 255, 0.4),
            0 0 15px rgba(0, 255, 170, 0.5),
            inset 0 0 5px rgba(255, 255, 255, 0.8);
        font-weight: 600;
        text-shadow: none;
    }}
    
    [data-testid="stSidebar"] .stButton button:active {{
        transform: translateY(1px) scale(0.98);
        box-shadow: 0 2px 10px rgba(0, 245, 255, 0.4);
    }}
    
    /* Sidebar Typography */
    [data-testid="stSidebar"] h1 {{
        color: {COLORS['foam_white']};
        font-weight: 800;
        font-size: 1.8rem;
        text-shadow: 0 0 15px rgba(0, 245, 255, 0.8), 2px 2px 5px rgba(0,0,0,0.8);
        letter-spacing: 2px;
        margin-bottom: 0.5rem;
    }}
    
    [data-testid="stSidebar"] .stMarkdown {{ color: {COLORS['light_aqua']}; }}
    [data-testid="stSidebar"] hr {{
        border-color: {COLORS['glass_border']};
        opacity: 0.5;
        margin: 1.5rem 0;
        box-shadow: 0 0 5px {COLORS['sky_blue']};
    }}
    
    /* Primary Buttons - Glowing Neon Wave */
    .stButton button[kind="primary"] {{
        background: linear-gradient(135deg, {COLORS['water_blue']} 0%, {COLORS['sky_blue']} 100%);
        color: {COLORS['deep_ocean']} !important;
        border: none;
        border-radius: 12px;
        padding: 14px 28px;
        font-weight: 700;
        font-size: 1rem;
        box-shadow: 0 0 20px rgba(0, 245, 255, 0.5), inset 0 2px 0 rgba(255, 255, 255, 0.4);
        transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
        text-transform: uppercase;
        letter-spacing: 1px;
        position: relative;
        overflow: hidden;
        animation: liquidGlow 4s infinite alternate;
    }}
    
    .stButton button[kind="primary"]:hover {{
        background: linear-gradient(135deg, {COLORS['sky_blue']} 0%, {COLORS['seafoam']} 100%);
        transform: translateY(-5px) scale(1.05);
        box-shadow: 0 15px 35px rgba(0, 245, 255, 0.6), 0 0 25px rgba(0, 255, 170, 0.6);
    }}
    
    /* Secondary Buttons - Deep Glass */
    .stButton button {{
        background: rgba(22, 22, 26, 0.5);
        color: {COLORS['light_aqua']} !important;
        border: 1px solid {COLORS['sky_blue']};
        border-radius: 10px;
        padding: 12px 24px;
        font-weight: 500;
        transition: all 0.4s ease;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }}
    
    .stButton button:hover {{
        background: rgba(139, 92, 246, 0.7);
        border-color: {COLORS['teal']};
        color: {COLORS['foam_white']} !important;
        box-shadow: 0 8px 25px rgba(0, 245, 255, 0.3);
        transform: translateY(-3px);
    }}
    
    /* Metrics - Hologram Text */
    [data-testid="stMetricValue"] {{
        font-size: 2.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, {COLORS['sky_blue']} 0%, {COLORS['teal']} 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-shadow: 0 0 30px rgba(0, 245, 255, 0.4);
        font-family: 'JetBrains Mono', monospace;
    }}
    
    [data-testid="stMetricLabel"] {{
        color: {COLORS['text_secondary']};
        font-weight: 600;
        font-size: 1rem;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }}
    
    /* Cards and Expanders - Frosted Abyss Glass */
    .stExpander {{
        background: rgba(22, 22, 26, 0.3);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid {COLORS['glass_border']};
        border-radius: 16px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.1);
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Expander content text visibility */
    .stExpander p, .stExpander div, .stExpander span, .stExpander li {{
        color: {COLORS['text_primary']} !important;
    }}
    
    .stExpander [data-testid="stMarkdownContainer"] {{
        color: {COLORS['text_primary']} !important;
    }}
    
    .stExpander:hover {{
        box-shadow: 0 15px 40px rgba(0, 245, 255, 0.2), inset 0 1px 0 rgba(255,255,255,0.2);
        transform: translateY(-5px);
        border-color: rgba(0, 245, 255, 0.6);
    }}
    
    /* DataFrames / Tables - Deep Sea Interface */
    .stDataFrame {{
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 8px 30px rgba(0,0,0,0.5), 0 0 15px rgba(0, 245, 255, 0.1);
        border: 1px solid {COLORS['glass_border']};
        background: rgba(10, 10, 14, 0.6) !important;
    }}
    
    table {{
        border-collapse: separate;
        border-spacing: 0;
        background: transparent !important;
        color: {COLORS['text_primary']} !important;
    }}
    
    thead tr {{
        background: linear-gradient(135deg, rgba(22, 22, 26, 0.9) 0%, rgba(139, 92, 246, 0.7) 100%) !important;
        color: {COLORS['foam_white']} !important;
        border-bottom: 2px solid {COLORS['sky_blue']} !important;
    }}
    
    thead th {{
        color: {COLORS['foam_white']} !important;
    }}
    
    tbody tr:nth-child(even) {{ background-color: rgba(22, 22, 26, 0.4) !important; }}
    tbody tr:nth-child(odd) {{ background-color: rgba(10, 10, 14, 0.6) !important; }}
    
    tbody tr:hover {{
        background-color: rgba(0, 245, 255, 0.15) !important;
        transition: background-color 0.3s ease;
        box-shadow: inset 0 0 10px rgba(0, 245, 255, 0.2);
    }}
    
    tbody td {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Success/Warning/Error Messages - Glowing Elements */
    .stSuccess {{
        background: linear-gradient(135deg, rgba(0, 255, 170, 0.05) 0%, rgba(0, 255, 170, 0.15) 100%);
        border: 1px solid rgba(0, 255, 170, 0.4);
        border-left: 6px solid {COLORS['success']};
        border-radius: 12px;
        color: {COLORS['text_primary']};
        box-shadow: 0 0 20px rgba(0, 255, 170, 0.15);
        backdrop-filter: blur(15px);
    }}
    .stWarning {{
        background: linear-gradient(135deg, rgba(255, 187, 0, 0.05) 0%, rgba(255, 187, 0, 0.15) 100%);
        border: 1px solid rgba(255, 187, 0, 0.4);
        border-left: 6px solid {COLORS['warning']};
        border-radius: 12px;
        color: {COLORS['text_primary']};
        box-shadow: 0 0 20px rgba(255, 187, 0, 0.15);
        backdrop-filter: blur(15px);
    }}
    .stError {{
        background: linear-gradient(135deg, rgba(255, 0, 85, 0.05) 0%, rgba(255, 0, 85, 0.15) 100%);
        border: 1px solid rgba(255, 0, 85, 0.4);
        border-left: 6px solid {COLORS['error']};
        border-radius: 12px;
        color: {COLORS['text_primary']};
        box-shadow: 0 0 20px rgba(255, 0, 85, 0.15);
        backdrop-filter: blur(15px);
    }}
    .stInfo {{
        background: linear-gradient(135deg, rgba(0, 245, 255, 0.05) 0%, rgba(0, 245, 255, 0.15) 100%);
        border: 1px solid rgba(0, 245, 255, 0.4);
        border-left: 6px solid {COLORS['info']};
        border-radius: 12px;
        color: {COLORS['text_primary']} !important;
        box-shadow: 0 0 20px rgba(0, 245, 255, 0.15);
        backdrop-filter: blur(15px);
    }}
    
    /* Fix container text visibility */
    .stInfo p, .stInfo div, .stInfo span, .stInfo li {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Fix all alert/message boxes */
    [data-testid="stNotification"], [data-testid="stAlert"] {{
        background: rgba(22, 22, 26, 0.8) !important;
        color: {COLORS['text_primary']} !important;
        border-radius: 12px;
    }}
    
    [data-testid="stNotification"] p, [data-testid="stAlert"] p,
    [data-testid="stNotification"] div, [data-testid="stAlert"] div {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Fix container backgrounds */
    div[data-testid="stVerticalBlock"] > div[style*="background"] {{
        background: rgba(22, 22, 26, 0.5) !important;
        border-radius: 12px;
        padding: 20px;
    }}
    
    /* Ensure all text in containers is visible */
    .element-container p, .element-container div, .element-container span {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Input Fields - Translucent Neon Bounds */
    .stTextInput input, .stNumberInput input, .stTextArea textarea {{
        background: rgba(10, 10, 14, 0.5) !important;
        color: {COLORS['text_primary']} !important;
        border: 1px solid {COLORS['glass_border']} !important;
        border-radius: 12px;
        padding: 14px;
        font-size: 1rem;
        transition: all 0.4s ease;
        box-shadow: inset 0 2px 5px rgba(0,0,0,0.5);
    }}
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {{
        border-color: {COLORS['sky_blue']} !important;
        box-shadow: 0 0 15px rgba(0, 245, 255, 0.4), inset 0 0 5px rgba(0, 245, 255, 0.2) !important;
        outline: none;
        background: rgba(22, 22, 26, 0.7) !important;
    }}
    
    /* Fix Selectbox and Multiselect Backgrounds */
    div[data-baseweb="select"] > div {{
        background-color: rgba(10, 10, 14, 0.8) !important;
        color: {COLORS['text_primary']} !important;
        border: 1px solid {COLORS['glass_border']} !important;
        border-radius: 12px !important;
        transition: all 0.4s ease;
    }}
    div[data-baseweb="select"] > div:hover {{
        border-color: {COLORS['sky_blue']} !important;
        box-shadow: 0 0 10px rgba(0, 245, 255, 0.2) !important;
    }}
    div[data-baseweb="select"] span, div[data-baseweb="select"] div {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Dropdown Menus */
    div[data-baseweb="popover"] ul {{
        background-color: rgba(10, 10, 14, 0.95) !important;
        border: 1px solid {COLORS['sky_blue']} !important;
        border-radius: 8px !important;
    }}
    div[data-baseweb="popover"] li {{
        color: {COLORS['text_primary']} !important;
    }}
    div[data-baseweb="popover"] li:hover {{
        background-color: rgba(0, 245, 255, 0.15) !important;
    }}
    
    /* FIX: Make the BROWSE FILES button text completely visible and properly positioned */
    [data-testid="stFileUploader"] section > button {{
        color: {COLORS['deep_ocean']} !important;
        background: linear-gradient(135deg, {COLORS['sky_blue']} 0%, {COLORS['light_aqua']} 100%) !important;
        font-weight: 800 !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 10px 24px !important;
        z-index: 10 !important;
        position: relative !important;
        box-shadow: 0 4px 15px rgba(0, 245, 255, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.6) !important;
        opacity: 1 !important;
        visibility: visible !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }}
    
    [data-testid="stFileUploader"] section > button:hover {{
        background: linear-gradient(135deg, {COLORS['light_aqua']} 0%, {COLORS['seafoam']} 100%) !important;
        box-shadow: 0 10px 25px rgba(0, 245, 255, 0.6) !important;
        transform: translateY(-2px) !important;
    }}
    
    /* File Uploader Container Fixes */
    [data-testid="stFileUploader"] {{
        background: linear-gradient(180deg, rgba(22, 22, 26, 0.2), rgba(10, 10, 14, 0.6)) !important;
        border: 2px dashed {COLORS['sky_blue']} !important;
        border-radius: 20px !important;
        padding: 2rem !important;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important;
        backdrop-filter: blur(15px) !important;
        -webkit-backdrop-filter: blur(15px) !important;
    }}
    [data-testid="stFileUploader"]:hover {{
        border-color: {COLORS['teal']} !important;
        background: linear-gradient(180deg, rgba(139, 92, 246, 0.2), rgba(22, 22, 26, 0.4)) !important;
        box-shadow: 0 10px 40px rgba(0, 245, 255, 0.2), inset 0 0 30px rgba(0, 255, 170, 0.1) !important;
        transform: scale(1.02) !important;
    }}
    
    /* File Uploader Dropzone Text Override */
    [data-testid="stFileUploaderDropzoneInstructions"] > div > span {{
        color: {COLORS['foam_white']} !important;
        font-weight: 700 !important;
        text-shadow: 0 0 10px rgba(0, 245, 255, 0.3) !important;
        font-size: 1.2rem !important;
        display: block !important;
        margin-bottom: 0.5rem !important;
    }}
    [data-testid="stFileUploaderDropzoneInstructions"] > div > small {{
        color: {COLORS['light_aqua']} !important;
        font-size: 0.95rem !important;
        opacity: 0.8 !important;
    }}
    
    /* Fix inside standard Uploader inner layout */
    section[data-testid="stFileUploaderDropzone"] {{
        background-color: transparent !important;
    }}
    
    /* Progress Bar - Flowing Neon Stream */
    .stProgress > div > div {{
        background: linear-gradient(90deg, {COLORS['water_blue']} 0%, {COLORS['sky_blue']} 50%, {COLORS['teal']} 100%);
        background-size: 200% 100%;
        animation: flowWater 1.5s linear infinite;
        border-radius: 10px;
        box-shadow: 0 0 15px {COLORS['sky_blue']};
    }}
    @keyframes flowWater {{
        0% {{ background-position: 100% 0%; }}
        100% {{ background-position: -100% 0%; }}
    }}
    
    /* Tabs - Floating Neon Labels */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 12px;
        background: rgba(22, 22, 26, 0.3);
        padding: 10px;
        border-radius: 16px;
        backdrop-filter: blur(20px);
        border: 1px solid {COLORS['glass_border']};
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 12px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.4s ease;
        color: {COLORS['text_muted']};
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {COLORS['ocean_blue']} 0%, {COLORS['water_blue']} 100%);
        color: {COLORS['text_primary']} !important;
        box-shadow: 0 5px 20px rgba(139, 92, 246, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.2);
        border: 1px solid {COLORS['sky_blue']};
    }}
    
    /* Download Button - Bioluminescent Warning Gradient */
    .stDownloadButton button {{
        background: linear-gradient(135deg, {COLORS['warning']} 0%, {COLORS['coral']} 100%) !important;
        color: {COLORS['deep_ocean']} !important;
        border: none;
        border-radius: 12px;
        padding: 14px 28px;
        font-weight: 800;
        box-shadow: 0 0 20px rgba(255, 0, 85, 0.4) !important;
        transition: all 0.4s ease !important;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }}
    .stDownloadButton button:hover {{
        transform: translateY(-5px) scale(1.03) !important;
        box-shadow: 0 10px 30px rgba(255, 187, 0, 0.6), 0 0 15px rgba(255, 0, 85, 0.6) !important;
    }}
    
    /* Headings - Glowing Text */
    h1, h2, h3 {{
        font-weight: 800;
        letter-spacing: -0.5px;
        text-transform: uppercase;
        color: {COLORS['foam_white']} !important;
    }}
    
    h1 {{
        font-size: 3rem;
        background: linear-gradient(135deg, {COLORS['light_aqua']} 0%, {COLORS['seafoam']} 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-shadow: 0 0 30px rgba(0, 245, 255, 0.3);
    }}
    h2 {{
        font-size: 2.2rem;
        color: {COLORS['sky_blue']} !important;
        text-shadow: 0 0 15px rgba(0, 245, 255, 0.2);
    }}
    h3 {{
        font-size: 1.5rem;
        color: {COLORS['teal']} !important;
    }}
    
    /* Smooth Pop In Animation */
    @keyframes deepPopIn {{
        0% {{ opacity: 0; transform: translateY(40px) scale(0.95); filter: blur(5px); }}
        100% {{ opacity: 1; transform: translateY(0) scale(1); filter: blur(0px); }}
    }}
    .element-container {{
        animation: deepPopIn 0.8s cubic-bezier(0.19, 1, 0.22, 1) forwards;
    }}
    
    /* Multiselect - Glowing Capsules */
    .stMultiSelect [data-baseweb="tag"] {{
        background: linear-gradient(135deg, {COLORS['water_blue']} 0%, {COLORS['sky_blue']} 100%);
        color: {COLORS['deep_ocean']};
        border-radius: 20px;
        padding: 6px 14px;
        font-weight: 700;
        box-shadow: 0 0 10px rgba(0, 245, 255, 0.4);
        border: 1px solid rgba(255, 255, 255, 0.3);
    }}
    
    /* Text improvements for dark mode */
    p, li, span, div {{
        font-size: 1.05rem;
        line-height: 1.7;
    }}
    
    .stMarkdown p, .stMarkdown li, .stMarkdown span {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Global text visibility fix - ensure all text is visible */
    p, span, div, li, label, td, th {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Force all text elements to be visible - aggressive fix */
    * {{
        color: {COLORS['text_primary']};
    }}
    
    /* But keep specific elements with their intended colors */
    h1, h2, h3 {{
        color: {COLORS['sky_blue']} !important;
    }}
    
    button {{
        color: {COLORS['deep_ocean']} !important;
    }}
    
    /* Specific fixes for common containers */
    [data-testid="stVerticalBlock"] p,
    [data-testid="stVerticalBlock"] span,
    [data-testid="stVerticalBlock"] div,
    [data-testid="stVerticalBlock"] label {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Fix for code blocks and pre elements */
    code, pre {{
        color: {COLORS['sky_blue']} !important;
        background: rgba(10, 10, 14, 0.8) !important;
    }}
    
    /* Subtexts */
    small, .stCaption {{
        color: {COLORS['text_muted']} !important;
    }}
    
    strong {{
        color: {COLORS['light_aqua']} !important;
        font-weight: 700;
        text-shadow: 0 0 10px rgba(0, 245, 255, 0.2);
    }}
    
    label {{
        color: {COLORS['text_primary']} !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px;
        font-size: 1rem !important;
    }}
    
    /* Specific fixes for form field labels */
    .stSelectbox label, .stTextInput label, .stNumberInput label, 
    .stTextArea label, .stDateInput label, .stTimeInput label {{
        color: {COLORS['text_primary']} !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        margin-bottom: 8px !important;
        display: block !important;
    }}
    
    /* Fix for column headers and section titles */
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, 
    .stMarkdown h4, .stMarkdown h5, .stMarkdown h6 {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Fix for any remaining invisible text */
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] span,
    [data-testid="stMarkdownContainer"] div {{
        color: {COLORS['text_primary']} !important;
    }}
    
    /* Ensure selection color is vibrant */
    ::selection {{
        background: {COLORS['sky_blue']};
        color: {COLORS['deep_ocean']};
    }}
    
    @media (max-width: 768px) {{
        h1 {{ font-size: 2.2rem; }}
        .stButton button {{ padding: 12px 20px; font-size: 0.9rem; }}
        [data-testid="stAppViewBlockContainer"] {{ padding: 1.5rem !important; margin-top: 1.5rem !important; }}
    }}
    </style>
    """, unsafe_allow_html=True)

    # Final theme override: matte cyber-intelligence report look.
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Michroma&family=Space+Mono:wght@400;700&display=swap');
    @import url('https://fonts.googleapis.com/icon?family=Material+Icons|Material+Icons+Outlined');
    @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,400,0,0&display=block');

    :root {
        --intel-bg: #070A0F;
        --intel-panel: #11161E;
        --intel-panel-2: #0D1118;
        --intel-border: #242C39;
        --intel-border-soft: #1A222E;
        --intel-text: #F8F3EA;
        --intel-muted: #91A0B8;
        --intel-faint: #617086;
        --intel-red: #F04348;
        --intel-amber: #FF9F0A;
        --intel-blue: #2F80FF;
        --intel-green: #21C16B;
        --intel-font: 'Space Mono', ui-monospace, SFMono-Regular, Consolas, monospace;
        --intel-display: 'Michroma', 'Space Mono', ui-monospace, monospace;
    }

    html, body, .stApp, [class*="css"] {
        font-family: var(--intel-font) !important;
        background: var(--intel-bg) !important;
        color: var(--intel-text) !important;
    }

    body * {
        font-family: var(--intel-font) !important;
        letter-spacing: 0 !important;
        text-shadow: none !important;
        animation: none !important;
    }

    .material-icons,
    .material-icons-outlined,
    .material-symbols-rounded,
    .material-symbols-outlined,
    [data-testid="stIconMaterial"],
    span[class*="material-icons"],
    span[class*="material-symbols"] {
        font-family: 'Material Symbols Rounded', 'Material Icons', 'Material Icons Outlined' !important;
        font-weight: normal !important;
        font-style: normal !important;
        font-size: 20px !important;
        line-height: 1 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 22px !important;
        min-width: 22px !important;
        max-width: 22px !important;
        height: 22px !important;
        overflow: hidden !important;
        white-space: nowrap !important;
        text-transform: none !important;
        letter-spacing: normal !important;
        word-wrap: normal !important;
        direction: ltr !important;
        -webkit-font-feature-settings: 'liga' !important;
        -webkit-font-smoothing: antialiased !important;
        font-feature-settings: 'liga' !important;
        color: var(--intel-muted) !important;
    }

    html body .stApp [data-testid="stExpander"] summary > span {
        font-family: var(--intel-font) !important;
        display: inline-flex !important;
        align-items: center !important;
        gap: 8px !important;
    }

    html body .stApp [data-testid="stExpander"] summary > span > span:first-child,
    html body .stApp details summary > span > span:first-child,
    html body .stApp [data-testid="stExpander"] summary > span > span:first-child *,
    html body .stApp details summary > span > span:first-child * {
        color: transparent !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        flex: 0 0 22px !important;
        width: 22px !important;
        min-width: 22px !important;
        max-width: 22px !important;
        height: 22px !important;
        margin-right: 4px !important;
        overflow: hidden !important;
        font-size: 0 !important;
        line-height: 1 !important;
        white-space: nowrap !important;
    }

    html body .stApp [data-testid="stExpander"] summary > span > span:first-child::before,
    html body .stApp details summary > span > span:first-child::before {
        content: "" !important;
        width: 0 !important;
        height: 0 !important;
        border-top: 7px solid transparent !important;
        border-bottom: 7px solid transparent !important;
        border-left: 10px solid var(--intel-amber) !important;
        display: block !important;
    }

    html body .stApp [data-testid="stExpander"] details[open] summary > span > span:first-child::before,
    html body .stApp details[open] summary > span > span:first-child::before {
        content: "" !important;
        border-left: 7px solid transparent !important;
        border-right: 7px solid transparent !important;
        border-top: 10px solid var(--intel-amber) !important;
        border-bottom: 0 !important;
    }

    button[data-testid="stBaseButton-headerNoPadding"] {
        color: transparent !important;
        font-size: 0 !important;
    }

    button[data-testid="stBaseButton-headerNoPadding"]::before {
        content: "\\2039" !important;
        color: var(--intel-muted) !important;
        font-family: Arial, sans-serif !important;
        font-size: 20px !important;
        line-height: 1 !important;
    }

    svg {
        flex: 0 0 auto !important;
    }

    header[data-testid="stHeader"], .stApp > header {
        background: transparent !important;
        height: 0 !important;
    }

    .stApp,
    [data-testid="stAppViewContainer"] {
        background: var(--intel-bg) !important;
    }

    [data-testid="stAppViewBlockContainer"] {
        max-width: 100% !important;
        padding: 32px !important;
        margin: 0 !important;
        background: transparent !important;
        border: 0 !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        backdrop-filter: none !important;
        -webkit-backdrop-filter: none !important;
    }

    [data-testid="stSidebar"] {
        background: #090C12 !important;
        border-right: 1px solid var(--intel-border) !important;
        box-shadow: none !important;
    }

    [data-testid="stSidebar"]::before {
        display: none !important;
    }

    [data-testid="stSidebar"] > div {
        background: transparent !important;
    }

    button[data-testid="stBaseButton-headerNoPadding"] {
        font-family: Arial, sans-serif !important;
        font-size: 0 !important;
        line-height: 1 !important;
        color: transparent !important;
        background: transparent !important;
        border: 0 !important;
        width: 28px !important;
        height: 28px !important;
        min-height: 28px !important;
        padding: 4px !important;
        overflow: hidden !important;
        white-space: nowrap !important;
        -webkit-font-feature-settings: 'liga' !important;
        font-feature-settings: 'liga' !important;
    }

    [class*="st-key-nav_upload"] button,
    [class*="st-key-nav_investment_fraud_top_accounts"] button,
    [class*="st-key-nav_top_10_suspect"] button,
    [class*="st-key-nav_e_fir_report"] button,
    [class*="st-key-nav_amount_matcher"] button,
    [class*="st-key-nav_transaction_matcher"] button,
    [class*="st-key-nav_view_database"] button {
        border-left: 3px solid var(--intel-red) !important;
    }

    [class*="st-key-nav_district_download"] button,
    [class*="st-key-nav_ifsc_pincode_split"] button,
    [class*="st-key-nav_bank_ack_pivot"] button,
    [class*="st-key-nav_call_notice_merge"] button,
    [class*="st-key-nav_bulk_mysql_import"] button,
    [class*="st-key-nav_ai_sql_assistant"] button {
        border-left: 3px solid var(--intel-amber) !important;
    }

    [class*="st-key-nav_districtwise"] button,
    [class*="st-key-nav_filter_by_entry_count"] button,
    [class*="st-key-nav_report_generator"] button,
    [class*="st-key-nav_automated_workflow"] button,
    [class*="st-key-nav_excel_merger"] button,
    [class*="st-key-nav_drop_call_finder"] button,
    [class*="st-key-nav_mo_finder"] button,
    [class*="st-key-nav_daily_report_district_split"] button,
    [class*="st-key-nav_ack_bank_consolidator"] button {
        border-left: 3px solid var(--intel-blue) !important;
    }

    [class*="st-key-nav_smart_district_split"] button,
    [class*="st-key-nav_filter_by_unique_ack"] button,
    [class*="st-key-nav_non_gujarat_filter"] button,
    [class*="st-key-nav_ack_list_pivot"] button,
    [class*="st-key-nav_column_selector"] button,
    [class*="st-key-nav_csv_fixer"] button,
    [class*="st-key-nav_disputed_amount_matcher"] button,
    [class*="st-key-nav_money_transfer_dispute"] button,
    [class*="st-key-nav_mysql_database_viewer"] button,
    [class*="st-key-nav_distinct_account_pivot"] button {
        border-left: 3px solid var(--intel-green) !important;
    }

    [class*="st-key-pin_nav_"] {
        --pin-accent: var(--intel-amber);
        --pin-accent-bg: rgba(255, 159, 10, 0.16);
    }

    [class*="st-key-pin_nav_upload"],
    [class*="st-key-pin_nav_investment_fraud_top_accounts"],
    [class*="st-key-pin_nav_top_10_suspect"],
    [class*="st-key-pin_nav_e_fir_report"],
    [class*="st-key-pin_nav_amount_matcher"],
    [class*="st-key-pin_nav_transaction_matcher"],
    [class*="st-key-pin_nav_view_database"] {
        --pin-accent: var(--intel-red);
        --pin-accent-bg: rgba(240, 67, 72, 0.16);
    }

    [class*="st-key-pin_nav_district_download"],
    [class*="st-key-pin_nav_ifsc_pincode_split"],
    [class*="st-key-pin_nav_bank_ack_pivot"],
    [class*="st-key-pin_nav_call_notice_merge"],
    [class*="st-key-pin_nav_bulk_mysql_import"],
    [class*="st-key-pin_nav_ai_sql_assistant"] {
        --pin-accent: var(--intel-amber);
        --pin-accent-bg: rgba(255, 159, 10, 0.16);
    }

    [class*="st-key-pin_nav_districtwise"],
    [class*="st-key-pin_nav_filter_by_entry_count"],
    [class*="st-key-pin_nav_report_generator"],
    [class*="st-key-pin_nav_automated_workflow"],
    [class*="st-key-pin_nav_excel_merger"],
    [class*="st-key-pin_nav_drop_call_finder"],
    [class*="st-key-pin_nav_mo_finder"],
    [class*="st-key-pin_nav_daily_report_district_split"],
    [class*="st-key-pin_nav_ack_bank_consolidator"] {
        --pin-accent: var(--intel-blue);
        --pin-accent-bg: rgba(47, 128, 255, 0.16);
    }

    [class*="st-key-pin_nav_smart_district_split"],
    [class*="st-key-pin_nav_filter_by_unique_ack"],
    [class*="st-key-pin_nav_non_gujarat_filter"],
    [class*="st-key-pin_nav_ack_list_pivot"],
    [class*="st-key-pin_nav_column_selector"],
    [class*="st-key-pin_nav_csv_fixer"],
    [class*="st-key-pin_nav_disputed_amount_matcher"],
    [class*="st-key-pin_nav_money_transfer_dispute"],
    [class*="st-key-pin_nav_mysql_database_viewer"],
    [class*="st-key-pin_nav_distinct_account_pivot"] {
        --pin-accent: var(--intel-green);
        --pin-accent-bg: rgba(33, 193, 107, 0.16);
    }

    [class*="st-key-pin_nav_"] button {
        border-left: 3px solid var(--pin-accent) !important;
        background: linear-gradient(90deg, var(--pin-accent-bg), var(--intel-panel) 64%) !important;
        color: var(--intel-text) !important;
    }

    .sidebar-logo {
        position: relative;
        margin: 12px 0 8px;
        padding: 0 0 4px;
    }

    .sidebar-brand {
        position: relative;
        display: inline-flex;
        align-items: baseline;
        color: var(--intel-text);
        font-family: var(--intel-display) !important;
        font-size: 1.42rem;
        font-weight: 400;
        line-height: 1.05;
        margin: 0;
        padding: 0 2px 0 0;
    }

    .sidebar-brand::before {
        content: "";
        position: absolute;
        left: -3px;
        top: 2px;
        width: 14px;
        height: 14px;
        border: 2px solid var(--intel-amber);
        border-right-color: var(--intel-blue);
        border-radius: 2px;
        transform: translateX(-2px);
    }

    .sidebar-brand span {
        font-family: var(--intel-display) !important;
    }

    .sidebar-brand .brand-data {
        color: var(--intel-text);
        padding-left: 2px;
    }

    .sidebar-brand .brand-lens {
        background: linear-gradient(90deg, var(--intel-blue), var(--intel-green) 58%, var(--intel-amber));
        -webkit-background-clip: text;
        background-clip: text;
        color: transparent !important;
        -webkit-text-fill-color: transparent !important;
    }

    .brand-accent-line {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr 1fr;
        gap: 3px;
        width: 136px;
        height: 3px;
        margin: 8px 0 7px;
    }

    .brand-accent-line span {
        display: block;
        border-radius: 999px;
    }

    .brand-accent-line span:nth-child(1) {
        background: var(--intel-red);
    }

    .brand-accent-line span:nth-child(2) {
        background: var(--intel-amber);
    }

    .brand-accent-line span:nth-child(3) {
        background: var(--intel-blue);
    }

    .brand-accent-line span:nth-child(4) {
        background: var(--intel-green);
    }

    .sidebar-subtitle,
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] caption {
        color: var(--intel-text) !important;
        font-size: 0.72rem !important;
        font-weight: 700 !important;
        text-transform: uppercase;
    }

    .sidebar-subtitle {
        display: inline-block;
        padding: 2px 6px;
        background: rgba(47, 128, 255, 0.10);
        border-left: 2px solid var(--intel-blue);
        border-right: 2px solid var(--intel-red);
        color: var(--intel-text) !important;
    }

    .sidebar-section-title {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        width: 100%;
        margin: 16px 0 8px;
        padding: 7px 9px 7px 11px;
        border-left: 3px solid var(--intel-blue);
        border-radius: 5px;
        background: linear-gradient(90deg, rgba(47, 128, 255, 0.18), rgba(17, 22, 30, 0.24) 68%);
        color: var(--intel-text) !important;
        font-family: var(--intel-font) !important;
        font-size: 0.74rem;
        font-weight: 800;
        letter-spacing: 0.02em;
        line-height: 1.2;
        text-transform: uppercase;
    }

    .sidebar-section-title::before {
        content: "";
        width: 6px;
        height: 6px;
        border-radius: 50%;
        background: var(--intel-green);
        box-shadow: 0 0 10px rgba(33, 193, 107, 0.58);
        flex: 0 0 auto;
    }

    .sidebar-section-title span {
        margin-left: auto;
        padding: 2px 7px;
        border: 1px solid rgba(255, 159, 10, 0.42);
        border-radius: 999px;
        background: rgba(255, 159, 10, 0.10);
        color: var(--intel-amber) !important;
        font-size: 0.68rem;
    }

    [data-testid="stSidebar"] hr,
    hr {
        border: 0 !important;
        border-top: 1px solid var(--intel-border-soft) !important;
        margin: 20px 0 !important;
        box-shadow: none !important;
        opacity: 1 !important;
    }

    h1, h2, h3,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3 {
        font-family: var(--intel-display) !important;
        position: relative;
        width: fit-content;
        max-width: 100%;
        font-weight: 400 !important;
        text-transform: none !important;
        background: none !important;
        -webkit-text-fill-color: currentColor !important;
        letter-spacing: 0 !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
    }

    h1,
    [data-testid="stMarkdownContainer"] h1 {
        font-size: 2rem !important;
        line-height: 1.05 !important;
        margin: 0 0 0.65rem !important;
        padding-left: 14px !important;
        color: #6EA6FF !important;
    }

    h1::before,
    [data-testid="stMarkdownContainer"] h1::before {
        content: "";
        position: absolute;
        left: 0;
        top: 0.12em;
        width: 3px;
        height: 0.92em;
        border-radius: 999px;
        background: linear-gradient(180deg, var(--intel-red), var(--intel-amber));
    }

    h1::after,
    [data-testid="stMarkdownContainer"] h1::after {
        content: "";
        display: block;
        width: min(190px, 62%);
        height: 2px;
        margin-top: 10px;
        border-radius: 999px;
        background: linear-gradient(90deg, var(--intel-red), var(--intel-amber), var(--intel-blue), var(--intel-green));
        opacity: 0.9;
    }

    h2,
    [data-testid="stMarkdownContainer"] h2 {
        font-size: 1.25rem !important;
        line-height: 1.2 !important;
        margin: 1.1rem 0 0.6rem !important;
        padding-left: 12px !important;
        color: #F2B35B !important;
        border-left: 2px solid var(--intel-amber);
    }

    h3,
    [data-testid="stMarkdownContainer"] h3 {
        font-size: 1rem !important;
        line-height: 1.25 !important;
        margin: 0.95rem 0 0.45rem !important;
        padding-left: 10px !important;
        color: #58D68D !important;
        border-left: 2px solid var(--intel-green);
    }

    p, li, label,
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li {
        color: var(--intel-muted) !important;
        font-size: 0.88rem !important;
        line-height: 1.55 !important;
    }

    strong {
        color: var(--intel-text) !important;
        font-weight: 700 !important;
    }

    small, .stCaption, [data-testid="stCaptionContainer"] {
        color: var(--intel-faint) !important;
        font-size: 0.72rem !important;
    }

    .intel-page-header {
        position: relative;
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 24px;
        padding: 2px 0 22px;
        border-bottom: 1px solid var(--intel-border-soft);
        box-shadow: inset 0 -2px 0 rgba(47, 128, 255, 0.16);
        margin-bottom: 26px;
    }

    .intel-kicker {
        display: inline-flex;
        align-items: center;
        gap: 10px;
        width: fit-content;
        font-size: 1.35rem;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .intel-page-header .intel-kicker {
        color: var(--intel-amber) !important;
    }

    .intel-kicker::before {
        content: "";
        width: 18px;
        height: 2px;
        border-radius: 999px;
        background: linear-gradient(90deg, var(--intel-red), var(--intel-amber));
        box-shadow: 0 0 12px rgba(255, 159, 10, 0.24);
    }

    .intel-page-header h1 {
        display: inline-block;
        color: var(--intel-text) !important;
        background: none !important;
        -webkit-text-fill-color: currentColor !important;
        font-family: var(--intel-display) !important;
        font-size: 2rem !important;
        line-height: 1.05 !important;
        padding-left: 0 !important;
        margin: 0 !important;
    }

    .intel-page-header h1::before {
        display: none !important;
    }

    .intel-page-header h1::after {
        content: "";
        display: block;
        width: 46%;
        min-width: 132px;
        height: 3px;
        margin-top: 11px;
        border-radius: 999px;
        background: linear-gradient(90deg, var(--intel-red), var(--intel-amber), var(--intel-blue), var(--intel-green));
        opacity: 0.95;
    }

    .intel-page-header h1 .accent {
        color: var(--intel-red) !important;
        -webkit-text-fill-color: var(--intel-red) !important;
        background: none !important;
        font-family: var(--intel-display) !important;
    }

    .intel-page-header h1 .title-base,
    .intel-page-header h1 .title-focus,
    .intel-page-header h1 .title-success,
    .intel-page-header h1 .title-hot,
    .datalens-global-title .title-base,
    .datalens-global-title .title-focus,
    .datalens-global-title .title-success,
    .datalens-global-title .title-hot {
        font-family: var(--intel-display) !important;
        font-size: inherit !important;
        line-height: inherit !important;
        background: none !important;
    }

    .intel-page-header h1 .title-base,
    .datalens-global-title .title-base {
        color: var(--intel-text) !important;
        -webkit-text-fill-color: var(--intel-text) !important;
    }

    .intel-page-header h1 .title-focus,
    .datalens-global-title .title-focus {
        color: var(--intel-blue) !important;
        -webkit-text-fill-color: var(--intel-blue) !important;
    }

    .intel-page-header h1 .title-success,
    .datalens-global-title .title-success {
        color: var(--intel-green) !important;
        -webkit-text-fill-color: var(--intel-green) !important;
    }

    .intel-page-header h1 .title-hot,
    .datalens-global-title .title-hot {
        color: var(--intel-red) !important;
        -webkit-text-fill-color: var(--intel-red) !important;
    }

    .intel-slogan {
        display: block;
        margin-top: 12px;
        color: #A2AEC2 !important;
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: none;
    }

    .intel-slogan .slogan-intel {
        color: var(--intel-red) !important;
    }

    .intel-slogan .slogan-trace {
        color: var(--intel-blue) !important;
    }

    .intel-slogan .slogan-report {
        color: var(--intel-green) !important;
    }

    .intel-slogan .slogan-sep {
        color: var(--intel-faint) !important;
    }

    .intel-page-header p {
        margin: 9px 0 0 !important;
        color: #7F93B4 !important;
        font-size: 0.74rem !important;
        text-transform: uppercase;
    }

    .intel-badge {
        border: 1px solid rgba(240, 67, 72, 0.45);
        background: rgba(240, 67, 72, 0.08);
        color: var(--intel-red);
        border-radius: 4px;
        padding: 7px 12px;
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
        white-space: nowrap;
    }

    .stButton button,
    [data-testid="stSidebar"] .stButton button {
        background: var(--intel-panel) !important;
        color: var(--intel-muted) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 6px !important;
        box-shadow: none !important;
        padding: 10px 14px !important;
        font-size: 0.76rem !important;
        font-weight: 700 !important;
        text-transform: none !important;
        transform: none !important;
        line-height: 1.25 !important;
        min-height: 38px !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
        transition: background-color 120ms ease, border-color 120ms ease, color 120ms ease !important;
    }

    [data-testid="stSidebar"] .stButton button {
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        text-align: left !important;
        color: #DDE8F6 !important;
        padding-left: 22px !important;
        padding-right: 12px !important;
    }

    [data-testid="stSidebar"] .stButton button p,
    [data-testid="stSidebar"] .stButton button span,
    [data-testid="stSidebar"] .stButton button div {
        color: #DDE8F6 !important;
        text-align: left !important;
        justify-content: flex-start !important;
    }

    .stButton button:hover,
    [data-testid="stSidebar"] .stButton button:hover {
        background: #151B24 !important;
        color: var(--intel-text) !important;
        border-color: var(--intel-red) !important;
        transform: none !important;
    }

    [data-testid="stSidebar"] .stButton button:hover,
    [data-testid="stSidebar"] .stButton button:focus,
    [data-testid="stSidebar"] .stButton button:active {
        color: var(--intel-text) !important;
    }

    [data-testid="stSidebar"] .stButton button:hover p,
    [data-testid="stSidebar"] .stButton button:hover span,
    [data-testid="stSidebar"] .stButton button:hover div,
    [data-testid="stSidebar"] .stButton button:focus p,
    [data-testid="stSidebar"] .stButton button:focus span,
    [data-testid="stSidebar"] .stButton button:focus div,
    [data-testid="stSidebar"] .stButton button:active p,
    [data-testid="stSidebar"] .stButton button:active span,
    [data-testid="stSidebar"] .stButton button:active div {
        color: var(--intel-text) !important;
    }

    [data-testid="stSidebar"] [class*="st-key-pin_current_page"] button {
        border-left: 3px solid var(--intel-green) !important;
        background: linear-gradient(90deg, rgba(33, 193, 107, 0.16), rgba(17, 22, 30, 0.96) 66%) !important;
        color: var(--intel-text) !important;
        box-shadow: inset 3px 0 0 var(--intel-green) !important;
    }

    [data-testid="stSidebar"] [class*="st-key-pin_current_page"] button:hover {
        border-color: var(--intel-green) !important;
        background: linear-gradient(90deg, rgba(33, 193, 107, 0.24), #151B24 66%) !important;
    }

    [data-testid="stSidebar"] [class*="st-key-unpin_current_page"] button {
        border-left: 3px solid var(--intel-red) !important;
        background: linear-gradient(90deg, rgba(240, 67, 72, 0.16), rgba(17, 22, 30, 0.96) 66%) !important;
        color: var(--intel-text) !important;
        box-shadow: inset 3px 0 0 var(--intel-red) !important;
    }

    [data-testid="stSidebar"] [class*="st-key-unpin_current_page"] button:hover {
        border-color: var(--intel-red) !important;
        background: linear-gradient(90deg, rgba(240, 67, 72, 0.24), #151B24 66%) !important;
    }

    [data-testid="stSidebar"] [class*="st-key-pin_current_page"] button p,
    [data-testid="stSidebar"] [class*="st-key-pin_current_page"] button span,
    [data-testid="stSidebar"] [class*="st-key-pin_current_page"] button div,
    [data-testid="stSidebar"] [class*="st-key-unpin_current_page"] button p,
    [data-testid="stSidebar"] [class*="st-key-unpin_current_page"] button span,
    [data-testid="stSidebar"] [class*="st-key-unpin_current_page"] button div {
        color: var(--intel-text) !important;
        font-weight: 800 !important;
    }

    [class*="st-key-show_pin_manager"] {
        margin: 8px 0 12px !important;
    }

    [class*="st-key-show_pin_manager"] label {
        width: 100% !important;
        padding: 8px 10px !important;
        border: 1px solid rgba(47, 128, 255, 0.32) !important;
        border-left: 3px solid var(--intel-blue) !important;
        border-radius: 6px !important;
        background: linear-gradient(90deg, rgba(47, 128, 255, 0.13), rgba(17, 22, 30, 0.40) 70%) !important;
    }

    [class*="st-key-show_pin_manager"] label p,
    [class*="st-key-show_pin_manager"] label span,
    [class*="st-key-show_pin_manager"] label div {
        color: #BFD7FF !important;
        font-family: var(--intel-font) !important;
        font-weight: 800 !important;
    }

    [class*="st-key-show_pin_manager"] [data-testid="stCheckbox"] {
        color: var(--intel-blue) !important;
    }

    [class*="st-key-nav_upload"] button,
    [class*="st-key-nav_investment_fraud_top_accounts"] button,
    [class*="st-key-nav_top_10_suspect"] button,
    [class*="st-key-nav_e_fir_report"] button,
    [class*="st-key-nav_amount_matcher"] button,
    [class*="st-key-nav_transaction_matcher"] button,
    [class*="st-key-nav_view_database"] button {
        border-left-color: var(--intel-red) !important;
        box-shadow: inset 3px 0 0 var(--intel-red) !important;
    }

    [class*="st-key-nav_upload"],
    [class*="st-key-nav_investment_fraud_top_accounts"],
    [class*="st-key-nav_top_10_suspect"],
    [class*="st-key-nav_e_fir_report"],
    [class*="st-key-nav_amount_matcher"],
    [class*="st-key-nav_transaction_matcher"],
    [class*="st-key-nav_view_database"] {
        border-left: 3px solid var(--intel-red) !important;
        border-radius: 6px !important;
    }

    [class*="st-key-nav_district_download"] button,
    [class*="st-key-nav_ifsc_pincode_split"] button,
    [class*="st-key-nav_bank_ack_pivot"] button,
    [class*="st-key-nav_call_notice_merge"] button,
    [class*="st-key-nav_bulk_mysql_import"] button,
    [class*="st-key-nav_ai_sql_assistant"] button {
        border-left-color: var(--intel-amber) !important;
        box-shadow: inset 3px 0 0 var(--intel-amber) !important;
    }

    [class*="st-key-nav_district_download"],
    [class*="st-key-nav_ifsc_pincode_split"],
    [class*="st-key-nav_bank_ack_pivot"],
    [class*="st-key-nav_call_notice_merge"],
    [class*="st-key-nav_bulk_mysql_import"],
    [class*="st-key-nav_ai_sql_assistant"] {
        border-left: 3px solid var(--intel-amber) !important;
        border-radius: 6px !important;
    }

    [class*="st-key-nav_districtwise"] button,
    [class*="st-key-nav_filter_by_entry_count"] button,
    [class*="st-key-nav_report_generator"] button,
    [class*="st-key-nav_automated_workflow"] button,
    [class*="st-key-nav_excel_merger"] button,
    [class*="st-key-nav_drop_call_finder"] button,
    [class*="st-key-nav_mo_finder"] button,
    [class*="st-key-nav_daily_report_district_split"] button,
    [class*="st-key-nav_ack_bank_consolidator"] button {
        border-left-color: var(--intel-blue) !important;
        box-shadow: inset 3px 0 0 var(--intel-blue) !important;
    }

    [class*="st-key-nav_districtwise"],
    [class*="st-key-nav_filter_by_entry_count"],
    [class*="st-key-nav_report_generator"],
    [class*="st-key-nav_automated_workflow"],
    [class*="st-key-nav_excel_merger"],
    [class*="st-key-nav_drop_call_finder"],
    [class*="st-key-nav_mo_finder"],
    [class*="st-key-nav_daily_report_district_split"],
    [class*="st-key-nav_ack_bank_consolidator"] {
        border-left: 3px solid var(--intel-blue) !important;
        border-radius: 6px !important;
    }

    [class*="st-key-nav_smart_district_split"] button,
    [class*="st-key-nav_filter_by_unique_ack"] button,
    [class*="st-key-nav_non_gujarat_filter"] button,
    [class*="st-key-nav_ack_list_pivot"] button,
    [class*="st-key-nav_column_selector"] button,
    [class*="st-key-nav_csv_fixer"] button,
    [class*="st-key-nav_disputed_amount_matcher"] button,
    [class*="st-key-nav_money_transfer_dispute"] button,
    [class*="st-key-nav_mysql_database_viewer"] button,
    [class*="st-key-nav_distinct_account_pivot"] button {
        border-left-color: var(--intel-green) !important;
        box-shadow: inset 3px 0 0 var(--intel-green) !important;
    }

    [class*="st-key-nav_smart_district_split"],
    [class*="st-key-nav_filter_by_unique_ack"],
    [class*="st-key-nav_non_gujarat_filter"],
    [class*="st-key-nav_ack_list_pivot"],
    [class*="st-key-nav_column_selector"],
    [class*="st-key-nav_csv_fixer"],
    [class*="st-key-nav_disputed_amount_matcher"],
    [class*="st-key-nav_money_transfer_dispute"],
    [class*="st-key-nav_mysql_database_viewer"],
    [class*="st-key-nav_distinct_account_pivot"] {
        border-left: 3px solid var(--intel-green) !important;
        border-radius: 6px !important;
    }

    [class*="st-key-pin_nav_"],
    [class*="st-key-pin_nav_"] button {
        border-left-color: var(--pin-accent) !important;
        box-shadow: inset 3px 0 0 var(--pin-accent) !important;
    }

    [class*="st-key-pin_nav_"] {
        border-left: 3px solid var(--pin-accent) !important;
        border-radius: 6px !important;
    }

    [class*="st-key-pin_nav_"] button p,
    [class*="st-key-pin_nav_"] button span,
    [class*="st-key-pin_nav_"] button div {
        color: var(--intel-text) !important;
    }

    .stButton button[kind="primary"] {
        background: var(--intel-red) !important;
        color: #090C12 !important;
        border-color: var(--intel-red) !important;
        box-shadow: none !important;
    }

    .stDownloadButton button {
        background: var(--intel-amber) !important;
        color: #090C12 !important;
        border: 1px solid var(--intel-amber) !important;
        border-radius: 6px !important;
        box-shadow: none !important;
        font-weight: 700 !important;
        text-transform: none !important;
        min-height: 38px !important;
        line-height: 1.25 !important;
        white-space: normal !important;
    }

    .stDownloadButton button *,
    .stDownloadButton button p,
    .stDownloadButton button span,
    .stDownloadButton button div {
        color: #090C12 !important;
    }

    div[data-testid="stMetric"],
    div[data-testid="metric-container"] {
        background: var(--intel-panel) !important;
        border: 1px solid var(--intel-border) !important;
        border-top: 3px solid var(--intel-red) !important;
        border-radius: 7px !important;
        padding: 16px !important;
        box-shadow: none !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetric"],
    [data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="metric-container"] {
        border-top-color: var(--intel-amber) !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(3) div[data-testid="stMetric"],
    [data-testid="stHorizontalBlock"] > div:nth-child(3) div[data-testid="metric-container"] {
        border-top-color: var(--intel-blue) !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(4) div[data-testid="stMetric"],
    [data-testid="stHorizontalBlock"] > div:nth-child(4) div[data-testid="metric-container"] {
        border-top-color: var(--intel-green) !important;
    }

    div[data-testid="stMetric"] label,
    [data-testid="stMetricLabel"] {
        color: var(--intel-faint) !important;
        font-size: 0.7rem !important;
        font-weight: 400 !important;
        text-transform: uppercase;
    }

    [data-testid="stMetricValue"] {
        color: var(--intel-red) !important;
        -webkit-text-fill-color: var(--intel-red) !important;
        background: none !important;
        font-family: var(--intel-display) !important;
        font-size: 2rem !important;
        font-weight: 400 !important;
        line-height: 1.1 !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(2) [data-testid="stMetricValue"] {
        color: var(--intel-amber) !important;
        -webkit-text-fill-color: var(--intel-amber) !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(3) [data-testid="stMetricValue"] {
        color: var(--intel-blue) !important;
        -webkit-text-fill-color: var(--intel-blue) !important;
    }

    [data-testid="stHorizontalBlock"] > div:nth-child(4) [data-testid="stMetricValue"] {
        color: var(--intel-green) !important;
        -webkit-text-fill-color: var(--intel-green) !important;
    }

    [data-testid="stMetricDelta"] {
        color: var(--intel-muted) !important;
        font-size: 0.74rem !important;
    }

    .stExpander,
    [data-testid="stForm"],
    [data-testid="stFileUploader"],
    div[data-testid="stAlert"],
    [data-testid="stNotification"] {
        background: var(--intel-panel) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 7px !important;
        box-shadow: none !important;
        backdrop-filter: none !important;
        -webkit-backdrop-filter: none !important;
    }

    .stExpander:hover {
        border-color: var(--intel-blue) !important;
        transform: none !important;
        box-shadow: none !important;
    }

    [data-testid="stFileUploader"] {
        background: linear-gradient(135deg, rgba(47, 128, 255, 0.10), rgba(17, 22, 30, 0.94)) !important;
        border-style: dashed !important;
        border-color: rgba(47, 128, 255, 0.45) !important;
        padding: 18px !important;
    }

    [data-testid="stFileUploader"] button,
    [data-testid="stFileUploader"] section > button {
        background: #151B24 !important;
        color: var(--intel-text) !important;
        border-radius: 6px !important;
        border: 1px solid var(--intel-red) !important;
        box-shadow: none !important;
        font-size: 0.75rem !important;
        font-weight: 700 !important;
        opacity: 1 !important;
        white-space: nowrap !important;
    }

    [data-testid="stFileUploader"] button *,
    [data-testid="stFileUploader"] section > button * {
        color: var(--intel-text) !important;
        opacity: 1 !important;
    }

    section[data-testid="stFileUploaderDropzone"] {
        background: transparent !important;
    }

    [data-testid="stFileUploaderDropzoneInstructions"] span {
        color: var(--intel-text) !important;
        font-size: 0.82rem !important;
        font-weight: 700 !important;
    }

    [data-testid="stFileUploaderDropzoneInstructions"] small {
        color: var(--intel-faint) !important;
    }

    .stTextInput input,
    .stNumberInput input,
    .stTextArea textarea,
    div[data-baseweb="select"] > div,
    [data-baseweb="input"] input {
        background: #0B0F15 !important;
        color: var(--intel-text) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 6px !important;
        box-shadow: none !important;
        font-size: 0.84rem !important;
    }

    .stTextInput input:focus,
    .stNumberInput input:focus,
    .stTextArea textarea:focus,
    div[data-baseweb="select"] > div:focus-within {
        border-color: var(--intel-blue) !important;
        box-shadow: none !important;
        outline: none !important;
    }

    div[data-baseweb="popover"] ul,
    div[data-baseweb="popover"] div {
        background: #0B0F15 !important;
        border-color: var(--intel-border) !important;
        color: var(--intel-text) !important;
    }

    div[data-baseweb="popover"] li:hover {
        background: #151B24 !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        background: var(--intel-panel) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 7px !important;
        padding: 5px !important;
        gap: 4px !important;
    }

    .stTabs [data-baseweb="tab"] {
        color: var(--intel-muted) !important;
        border-radius: 5px !important;
        padding: 8px 12px !important;
        font-size: 0.78rem !important;
        font-weight: 700 !important;
    }

    .stTabs [aria-selected="true"] {
        background: var(--intel-red) !important;
        color: #090C12 !important;
        border: 0 !important;
        box-shadow: none !important;
    }

    .stDataFrame,
    [data-testid="stTable"] {
        background: var(--intel-panel) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 7px !important;
        box-shadow: none !important;
        overflow: hidden !important;
    }

    table {
        background: var(--intel-panel) !important;
        color: var(--intel-text) !important;
        border-collapse: collapse !important;
    }

    thead tr {
        background: #0B0F15 !important;
        border-bottom: 1px solid var(--intel-border) !important;
    }

    thead th {
        color: var(--intel-faint) !important;
        font-size: 0.74rem !important;
        text-transform: uppercase;
    }

    tbody tr:nth-child(even),
    tbody tr:nth-child(odd) {
        background: var(--intel-panel) !important;
    }

    tbody tr {
        border-bottom: 1px solid var(--intel-border-soft) !important;
    }

    tbody tr:hover {
        background: #151B24 !important;
        box-shadow: none !important;
    }

    tbody td {
        color: var(--intel-text) !important;
        font-size: 0.8rem !important;
    }

    .stProgress > div {
        background: #1A222E !important;
    }

    .stProgress > div > div {
        background: var(--intel-red) !important;
        border-radius: 2px !important;
        box-shadow: none !important;
    }

    .stSuccess {
        border-left: 4px solid var(--intel-green) !important;
    }

    .stWarning {
        border-left: 4px solid var(--intel-amber) !important;
    }

    .stError {
        border-left: 4px solid var(--intel-red) !important;
    }

    .stInfo {
        border-left: 4px solid var(--intel-blue) !important;
    }

    code, pre {
        background: #0B0F15 !important;
        color: var(--intel-amber) !important;
        border: 1px solid var(--intel-border) !important;
        border-radius: 4px !important;
    }

    ::selection {
        background: var(--intel-red);
        color: #090C12;
    }

    ::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }

    ::-webkit-scrollbar-track {
        background: var(--intel-bg);
    }

    ::-webkit-scrollbar-thumb {
        background: #2A3341;
        border: 2px solid var(--intel-bg);
        border-radius: 8px;
    }

    ::-webkit-scrollbar-thumb:hover {
        background: var(--intel-red);
        box-shadow: none;
    }

    #sidebar-resizer:hover,
    #sidebar-resizer.active {
        background: var(--intel-red) !important;
        box-shadow: none !important;
    }

    @media (max-width: 768px) {
        [data-testid="stAppViewBlockContainer"] {
            padding: 18px !important;
        }

        .intel-page-header {
            align-items: flex-start;
            flex-direction: column;
            gap: 12px;
        }

        .intel-page-header h1,
        h1 {
            font-size: 1.8rem !important;
        }

        .intel-kicker {
            font-size: 1.05rem;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    _install_heading_patches()


def _clean_display_text(value):
    """Keep headings crisp by removing emoji/mojibake prefixes from legacy labels."""
    ascii_text = ''.join(ch if 32 <= ord(ch) < 127 else ' ' for ch in str(value))
    return ' '.join(ascii_text.split())


def _get_time_greeting():
    """Return the fixed DataLens greeting."""
    return "Hi, let's make the numbers talk."


def _render_title_html(title):
    """Render page titles with restrained theme accents."""
    words = _clean_display_text(title).split()
    if not words:
        return ""

    rendered = []
    last_index = len(words) - 1
    for index, word in enumerate(words):
        if word.upper() == "MO":
            class_name = "title-hot"
        elif index == 0:
            class_name = "title-base"
        elif index == last_index:
            class_name = "title-success"
        else:
            class_name = "title-focus"
        rendered.append(f'<span class="{class_name}">{escape(word)}</span>')
    return " ".join(rendered)


def _install_heading_patches():
    """Apply the DataLens heading system to every page module."""
    if getattr(st, "_datalens_heading_patch_installed", False):
        return

    def themed_title(body, *args, **kwargs):
        text = _clean_display_text(body)
        if not text:
            return None
        st.markdown(
            f'<h1 class="datalens-global-title">{_render_title_html(text)}</h1>',
            unsafe_allow_html=True
        )
        return None

    def themed_header(body, *args, **kwargs):
        text = _clean_display_text(body)
        if not text:
            return None
        st.markdown(
            f'<h2 class="datalens-section-title">{escape(text)}</h2>',
            unsafe_allow_html=True
        )
        return None

    def themed_subheader(body, *args, **kwargs):
        text = _clean_display_text(body)
        if not text:
            return None
        st.markdown(
            f'<h3 class="datalens-subsection-title">{escape(text)}</h3>',
            unsafe_allow_html=True
        )
        return None

    st._datalens_original_title = st.title
    st._datalens_original_header = st.header
    st._datalens_original_subheader = st.subheader
    st.title = themed_title
    st.header = themed_header
    st.subheader = themed_subheader
    st._datalens_heading_patch_installed = True


def render_page_header_with_info(page_key):
    """Render page header with information button"""
    if page_key not in PAGE_INFO:
        return
    
    info = PAGE_INFO[page_key]
    title = _clean_display_text(info['title'])
    description = _clean_display_text(info['description'])
    greeting = _get_time_greeting()
    title_html = _render_title_html(title)
    slogan = "Cyber fraud intelligence for faster review and clear reports."
    
    # Create header with info button
    col1, col2 = st.columns([0.88, 0.12])
    
    with col1:
        st.markdown(f"""
        <section class="intel-page-header">
            <div>
                <div class="intel-kicker">{escape(greeting)}</div>
                <h1>{title_html}</h1>
                <div class="intel-slogan">{escape(slogan)}</div>
                <p>{escape(description)}</p>
            </div>
            <div class="intel-badge">Confidential Intelligence</div>
        </section>
        """, unsafe_allow_html=True)
    
    with col2:
        # Info button
        if st.button("INFO", key=f"info_btn_{page_key}", help="Click for detailed information"):
            st.session_state[f'show_info_{page_key}'] = not st.session_state.get(f'show_info_{page_key}', False)
    
    # Show detailed info if button clicked
    if st.session_state.get(f'show_info_{page_key}', False):
        with st.expander("Detailed Information", expanded=True):
            st.markdown(info['details'])
            if st.button("Close", key=f"close_info_{page_key}"):
                st.session_state[f'show_info_{page_key}'] = False
                st.rerun()
    
    st.markdown("")


def render_feature_card(title, description, icon="📊"):
    """Render a feature card with icon"""
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, rgba(22, 22, 26, 0.6) 0%, rgba(10, 10, 14, 0.8) 100%); 
                border-radius: 16px; padding: 24px; 
                box-shadow: 0 10px 30px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.1); 
                margin-bottom: 24px;
                backdrop-filter: blur(15px); -webkit-backdrop-filter: blur(15px); border: 1px solid {COLORS['glass_border']};
                transition: all 0.4s ease;
                cursor: pointer;"
         onmouseover="this.style.transform='translateY(-5px)'; this.style.boxShadow='0 15px 40px rgba(0, 245, 255, 0.3), inset 0 1px 0 rgba(255,255,255,0.2)'; this.style.borderColor='{COLORS['teal']}';"
         onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 10px 30px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.1)'; this.style.borderColor='{COLORS['glass_border']}';">
        <h3 style="color: {COLORS['light_aqua']}; font-size: 1.4rem; margin-bottom: 12px; text-shadow: 0 0 10px rgba(0, 245, 255, 0.4);">{icon} {title}</h3>
        <p style="color: {COLORS['text_muted']}; font-size: 1.05rem; line-height: 1.6; margin: 0;">{description}</p>
    </div>
    """, unsafe_allow_html=True)


def render_feature_card(title, description, icon=""):
    """Render a feature card using the cyber-intelligence report theme."""
    title_text = escape(_clean_display_text(title))
    description_text = escape(_clean_display_text(description))
    icon_text = escape(_clean_display_text(icon))
    prefix = f"{icon_text} " if icon_text else ""
    st.markdown(f"""
    <div style="background: {COLORS['background_card']};
                border: 1px solid {COLORS['border']};
                border-top: 3px solid {COLORS['light_aqua']};
                border-radius: 7px;
                padding: 18px;
                margin-bottom: 18px;
                box-shadow: none;">
        <h3 style="color: {COLORS['text_primary']} !important;
                   font-size: 0.95rem;
                   margin: 0 0 8px 0;
                   font-family: 'Space Mono', Consolas, monospace !important;">{prefix}{title_text}</h3>
        <p style="color: {COLORS['text_secondary']} !important;
                  font-size: 0.82rem;
                  line-height: 1.55;
                  margin: 0;">{description_text}</p>
    </div>
    """, unsafe_allow_html=True)
