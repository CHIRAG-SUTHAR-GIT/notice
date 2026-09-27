"""
Call Notice Data Merge - Match records from two files based on mobile numbers.
Calculates time differences between call date and entry date.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import re


def normalize_mobile(mobile):
    """
    Extract last 10 digits from mobile number.
    Handles: country code 91, scientific notation, floats, strings with spaces/dashes
    """
    if pd.isna(mobile):
        return None
    
    # Convert to string
    mobile_str = str(mobile).strip()
    
    # Handle scientific notation (e.g., 9.19876543210e+11)
    if 'e' in mobile_str.lower():
        try:
            # Use Decimal for precision with large numbers
            from decimal import Decimal, ROUND_DOWN
            mobile_str = str(int(Decimal(mobile_str).to_integral_value(rounding=ROUND_DOWN)))
        except:
            try:
                mobile_str = format(int(float(mobile_str)), 'd')
            except:
                pass
    
    # Remove .0 suffix if present (from float conversion)
    if '.' in mobile_str:
        try:
            # Handle float strings like "919876543210.0"
            mobile_str = str(int(float(mobile_str)))
        except:
            mobile_str = mobile_str.split('.')[0]
    
    # Remove any non-digit characters (spaces, dashes, brackets, etc.)
    mobile_str = re.sub(r'\D', '', mobile_str)
    
    # Extract last 10 digits only
    if len(mobile_str) >= 10:
        return mobile_str[-10:]
    
    # If less than 10 digits, return None (invalid)
    return None


def parse_datetime(dt_value):
    """
    Parse datetime from various formats with high accuracy.
    Returns pandas Timestamp for precise calculations.
    """
    if pd.isna(dt_value):
        return pd.NaT
    
    # If already datetime/Timestamp, convert to pandas Timestamp
    if isinstance(dt_value, pd.Timestamp):
        return dt_value
    if isinstance(dt_value, datetime):
        return pd.Timestamp(dt_value)
    
    # Try pandas to_datetime first (handles many formats automatically)
    try:
        parsed = pd.to_datetime(dt_value, dayfirst=True)
        if pd.notna(parsed):
            return parsed
    except:
        pass
    
    # Try parsing string formats manually
    dt_str = str(dt_value).strip()
    formats = [
        '%d-%m-%Y %H:%M:%S',
        '%d-%m-%Y %H:%M',
        '%d/%m/%Y %H:%M:%S',
        '%d/%m/%Y %H:%M',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M:%S.%f',
        '%Y-%m-%d %H:%M',
        '%Y/%m/%d %H:%M:%S',
        '%Y/%m/%d %H:%M',
        '%m/%d/%Y %H:%M:%S',
        '%m/%d/%Y %H:%M',
        '%d-%m-%Y',
        '%d/%m/%Y',
        '%Y-%m-%d',
    ]
    
    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(dt_str, fmt))
        except ValueError:
            continue
    
    return pd.NaT


def calculate_time_difference(entry_date, call_date):
    """
    Calculate time difference: Entry Date - Call Date
    Returns timedelta object for accurate calculations.
    """
    if pd.isna(entry_date) or pd.isna(call_date):
        return None
    
    try:
        diff = entry_date - call_date
        return diff
    except:
        return None


def format_time_difference(td):
    """
    Format timedelta as H:MM:SS with proper handling of negative values.
    """
    if td is None or pd.isna(td):
        return None
    
    try:
        total_seconds = td.total_seconds()
        is_negative = total_seconds < 0
        total_seconds = abs(int(total_seconds))
        
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        
        result = f"{hours}:{minutes:02d}:{seconds:02d}"
        return f"-{result}" if is_negative else result
    except:
        return None


def format_datetime_output(dt):
    """Format datetime as DD-MM-YYYY HH:MM."""
    if pd.isna(dt):
        return None
    try:
        if isinstance(dt, (pd.Timestamp, datetime)):
            return dt.strftime('%d-%m-%Y %H:%M')
        return str(dt)
    except:
        return str(dt)


def validate_mobile(mobile):
    """Validate that mobile number is exactly 10 digits."""
    if mobile is None or pd.isna(mobile):
        return False
    mobile_str = str(mobile)
    return len(mobile_str) == 10 and mobile_str.isdigit()


def render_call_notice_merge_page():
    """Render the Call Notice Data Merge page."""
    
    st.title("📊 Call Notice Data Merge")
    st.markdown("Match records from two Excel files based on mobile numbers and calculate time differences.")
    
    # Initialize session state for this page
    if 'cnm_file1_df' not in st.session_state:
        st.session_state.cnm_file1_df = None
    if 'cnm_file2_df' not in st.session_state:
        st.session_state.cnm_file2_df = None
    if 'cnm_matched_df' not in st.session_state:
        st.session_state.cnm_matched_df = None
    
    # File Upload Section
    st.header("📁 Upload Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("File 1: Call Data")
        file1 = st.file_uploader("Upload Call Data file", type=['xlsx', 'xls', 'csv'], key='cnm_file1')
        
        if file1:
            try:
                if file1.name.endswith('.csv'):
                    st.session_state.cnm_file1_df = pd.read_csv(file1)
                else:
                    st.session_state.cnm_file1_df = pd.read_excel(file1)
                st.success(f"✅ Loaded {len(st.session_state.cnm_file1_df)} records")
                
                with st.expander("Preview Call Data", expanded=True):
                    st.dataframe(st.session_state.cnm_file1_df.head(10), use_container_width=True)
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
    
    with col2:
        st.subheader("File 2: Entry Data")
        file2 = st.file_uploader("Upload Entry Data file", type=['xlsx', 'xls', 'csv'], key='cnm_file2')
        
        if file2:
            try:
                if file2.name.endswith('.csv'):
                    st.session_state.cnm_file2_df = pd.read_csv(file2)
                else:
                    st.session_state.cnm_file2_df = pd.read_excel(file2)
                st.success(f"✅ Loaded {len(st.session_state.cnm_file2_df)} records")
                
                with st.expander("Preview Entry Data", expanded=True):
                    st.dataframe(st.session_state.cnm_file2_df.head(10), use_container_width=True)
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
    
    # Column Selection Section
    if st.session_state.cnm_file1_df is not None and st.session_state.cnm_file2_df is not None:
        st.header("🔧 Column Selection")
        
        col1, col2 = st.columns(2)
        
        file1_columns = st.session_state.cnm_file1_df.columns.tolist()
        file2_columns = st.session_state.cnm_file2_df.columns.tolist()
        
        with col1:
            st.subheader("File 1 Columns (Call Data)")
            phone_col_file1 = st.selectbox(
                "Select Phone Number Column",
                options=file1_columns,
                key='cnm_phone_col_file1'
            )
            call_date_col = st.selectbox(
                "Select Call Date Column",
                options=file1_columns,
                key='cnm_call_date_col'
            )
        
        with col2:
            st.subheader("File 2 Columns (Entry Data)")
            ack_col = st.selectbox(
                "Select Acknowledgement Number Column",
                options=file2_columns,
                key='cnm_ack_col'
            )
            phone_col_file2 = st.selectbox(
                "Select Mobile Number Column",
                options=file2_columns,
                key='cnm_phone_col_file2'
            )
            entry_date_col = st.selectbox(
                "Select Entry Date Column",
                options=file2_columns,
                key='cnm_entry_date_col'
            )
        
        # Process and Match Button
        st.header("🔄 Process Data")
        
        # Debug preview section
        with st.expander("🔍 Preview Normalized Data (Verify Before Matching)", expanded=False):
            col_debug1, col_debug2 = st.columns(2)
            
            with col_debug1:
                st.write("**File 1 - Sample Mobile Numbers & Dates:**")
                sample_df1 = st.session_state.cnm_file1_df[[phone_col_file1, call_date_col]].head(10).copy()
                sample_df1['Normalized Mobile'] = sample_df1[phone_col_file1].apply(normalize_mobile)
                sample_df1['Parsed Date'] = sample_df1[call_date_col].apply(parse_datetime)
                sample_df1['Parsed Date'] = sample_df1['Parsed Date'].apply(format_datetime_output)
                st.dataframe(sample_df1, use_container_width=True)
            
            with col_debug2:
                st.write("**File 2 - Sample Mobile Numbers & Dates:**")
                sample_df2 = st.session_state.cnm_file2_df[[phone_col_file2, entry_date_col]].head(10).copy()
                sample_df2['Normalized Mobile'] = sample_df2[phone_col_file2].apply(normalize_mobile)
                sample_df2['Parsed Date'] = sample_df2[entry_date_col].apply(parse_datetime)
                sample_df2['Parsed Date'] = sample_df2['Parsed Date'].apply(format_datetime_output)
                st.dataframe(sample_df2, use_container_width=True)
        
        if st.button("Match Records", type="primary", use_container_width=True, key="cnm_match_btn"):
            with st.spinner("Processing data..."):
                try:
                    # Create working copies
                    df1 = st.session_state.cnm_file1_df.copy()
                    df2 = st.session_state.cnm_file2_df.copy()
                    
                    # Step 1: Normalize mobile numbers
                    df1['normalized_mobile'] = df1[phone_col_file1].apply(normalize_mobile)
                    df2['normalized_mobile'] = df2[phone_col_file2].apply(normalize_mobile)
                    
                    # Track data quality issues
                    warnings = []
                    
                    # Check for invalid mobile numbers
                    invalid_mobile_file1 = df1[~df1['normalized_mobile'].apply(validate_mobile)]
                    invalid_mobile_file2 = df2[~df2['normalized_mobile'].apply(validate_mobile)]
                    
                    if len(invalid_mobile_file1) > 0:
                        warnings.append(f"⚠️ File 1: {len(invalid_mobile_file1)} records with invalid mobile numbers (excluded)")
                    if len(invalid_mobile_file2) > 0:
                        warnings.append(f"⚠️ File 2: {len(invalid_mobile_file2)} records with invalid mobile numbers (excluded)")
                    
                    # Step 2: Filter valid mobile numbers only
                    df1_valid = df1[df1['normalized_mobile'].apply(validate_mobile)].copy()
                    df2_valid = df2[df2['normalized_mobile'].apply(validate_mobile)].copy()
                    
                    # Step 3: Parse dates
                    df1_valid['parsed_call_date'] = df1_valid[call_date_col].apply(parse_datetime)
                    df2_valid['parsed_entry_date'] = df2_valid[entry_date_col].apply(parse_datetime)
                    
                    # Check for unparseable dates
                    unparsed_call = df1_valid['parsed_call_date'].isna().sum()
                    unparsed_entry = df2_valid['parsed_entry_date'].isna().sum()
                    
                    if unparsed_call > 0:
                        warnings.append(f"⚠️ File 1: {unparsed_call} records with unparseable call dates")
                    if unparsed_entry > 0:
                        warnings.append(f"⚠️ File 2: {unparsed_entry} records with unparseable entry dates")
                    
                    # Step 4: Prepare data for merge
                    # KEEP ONLY FIRST CALL for each mobile number (remove duplicates)
                    df1_for_merge = df1_valid[['normalized_mobile', 'parsed_call_date']].copy()
                    df1_for_merge = df1_for_merge.drop_duplicates(subset='normalized_mobile', keep='first')
                    df1_for_merge = df1_for_merge.reset_index(drop=True)
                    
                    # Count duplicates removed
                    duplicates_removed = len(df1_valid) - len(df1_for_merge)
                    if duplicates_removed > 0:
                        warnings.append(f"ℹ️ File 1: {duplicates_removed} duplicate calls removed (kept first call only)")
                    
                    df2_for_merge = df2_valid[['normalized_mobile', ack_col, 'parsed_entry_date']].copy()
                    df2_for_merge = df2_for_merge.rename(columns={ack_col: 'ack_no'})
                    # Also remove duplicates from File 2 if any (keep first entry)
                    df2_duplicates = len(df2_for_merge) - len(df2_for_merge.drop_duplicates(subset='normalized_mobile', keep='first'))
                    df2_for_merge = df2_for_merge.drop_duplicates(subset='normalized_mobile', keep='first')
                    df2_for_merge = df2_for_merge.reset_index(drop=True)
                    
                    if df2_duplicates > 0:
                        warnings.append(f"ℹ️ File 2: {df2_duplicates} duplicate entries removed (kept first entry only)")
                    
                    # Step 5: Perform INNER JOIN on normalized mobile numbers
                    merged = pd.merge(
                        df1_for_merge,
                        df2_for_merge,
                        on='normalized_mobile',
                        how='inner'
                    )
                    
                    # Step 6: Calculate time difference with precision
                    merged['time_diff'] = merged.apply(
                        lambda row: calculate_time_difference(row['parsed_entry_date'], row['parsed_call_date']),
                        axis=1
                    )
                    
                    # Step 6.5: Filter out records with negative time differences
                    original_count = len(merged)
                    # Keep only records where time_diff is positive or zero (not negative)
                    merged = merged[merged['time_diff'].apply(lambda x: x is not None and pd.notna(x) and x.total_seconds() >= 0)].copy()
                    filtered_count = len(merged)
                    negative_count = original_count - filtered_count
                    
                    if negative_count > 0:
                        st.warning(f"⚠️ Removed {negative_count} record(s) with negative time difference (Call Date after Entry Date)")
                    
                    # Step 7: Create final output dataframe
                    output_df = pd.DataFrame({
                        'Sr. No.': range(1, len(merged) + 1),
                        'Acknowledgement No.': merged['ack_no'],
                        'Mobile Number': merged['normalized_mobile'],
                        'Call Date': merged['parsed_call_date'].apply(format_datetime_output),
                        'Entry Date': merged['parsed_entry_date'].apply(format_datetime_output),
                        'Time Difference': merged['time_diff'].apply(format_time_difference)
                    })
                    
                    # Calculate average time difference (only from positive time differences)
                    valid_time_diffs = merged['time_diff'].dropna()
                    if len(valid_time_diffs) > 0:
                        avg_seconds = valid_time_diffs.apply(lambda x: x.total_seconds()).mean()
                        avg_timedelta = timedelta(seconds=avg_seconds)
                        avg_formatted = format_time_difference(avg_timedelta)
                        
                        # Add average row at bottom
                        avg_row = pd.DataFrame({
                            'Sr. No.': [''],
                            'Acknowledgement No.': [''],
                            'Mobile Number': [''],
                            'Call Date': [''],
                            'Entry Date': ['AVERAGE'],
                            'Time Difference': [avg_formatted]
                        })
                        output_df = pd.concat([output_df, avg_row], ignore_index=True)
                    
                    st.session_state.cnm_matched_df = output_df
                    
                    # Display warnings
                    if warnings:
                        st.warning("Data Quality Issues Detected:")
                        for w in warnings:
                            st.write(w)
                    
                    # Display summary statistics
                    st.header("📈 Summary Statistics")
                    
                    stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
                    
                    with stats_col1:
                        st.metric("File 1 Total Records", len(st.session_state.cnm_file1_df))
                    with stats_col2:
                        st.metric("File 2 Total Records", len(st.session_state.cnm_file2_df))
                    with stats_col3:
                        st.metric("Matched Records", len(output_df) - 1 if len(output_df) > 0 else 0)  # -1 for average row
                    with stats_col4:
                        # Calculate unique mobiles that didn't match
                        file1_mobiles = set(df1_valid['normalized_mobile'].dropna())
                        file2_mobiles = set(df2_valid['normalized_mobile'].dropna())
                        unmatched_f1 = len(file1_mobiles - file2_mobiles)
                        unmatched_f2 = len(file2_mobiles - file1_mobiles)
                        st.metric("Unmatched Unique Mobiles", f"{unmatched_f1} / {unmatched_f2}")
                    
                    st.success(f"✅ Successfully matched {len(output_df) - 1} records!")
                    
                    # Verification section
                    with st.expander("🔬 Verify Sample Calculations", expanded=True):
                        if len(output_df) > 0:
                            st.write("**First 5 matched records with calculation verification:**")
                            verify_df = merged.head(5).copy()
                            verify_df['Call Date Raw'] = verify_df['parsed_call_date']
                            verify_df['Entry Date Raw'] = verify_df['parsed_entry_date']
                            verify_df['Time Diff (seconds)'] = verify_df['time_diff'].apply(
                                lambda x: x.total_seconds() if x is not None and pd.notna(x) else None
                            )
                            verify_df['Time Diff Formatted'] = verify_df['time_diff'].apply(format_time_difference)
                            st.dataframe(verify_df[['normalized_mobile', 'Call Date Raw', 'Entry Date Raw', 
                                                    'Time Diff (seconds)', 'Time Diff Formatted']], 
                                        use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ Error during processing: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        
        # Display matched results and download option
        if st.session_state.cnm_matched_df is not None and len(st.session_state.cnm_matched_df) > 0:
            st.header("📋 Matched Results Preview")
            st.dataframe(st.session_state.cnm_matched_df, use_container_width=True)
            
            # Download buttons
            st.subheader("📥 Download Results")
            download_col1, download_col2 = st.columns(2)
            
            with download_col1:
                # Create Excel file for download using openpyxl
                excel_output = BytesIO()
                with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                    st.session_state.cnm_matched_df.to_excel(writer, index=False, sheet_name='Matched Records')
                excel_output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_output,
                    file_name=f"matched_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            
            with download_col2:
                # Create CSV file for download
                csv_output = st.session_state.cnm_matched_df.to_csv(index=False)
                
                st.download_button(
                    label="📥 Download CSV (.csv)",
                    data=csv_output,
                    file_name=f"matched_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    type="primary",
                    use_container_width=True
                )
        elif st.session_state.cnm_matched_df is not None and len(st.session_state.cnm_matched_df) == 0:
            st.warning("⚠️ No matching records found between the two files.")
    
    # Footer
    st.markdown("---")
    st.markdown("*Call Notice Data Merge - Match records based on mobile numbers and calculate time differences*")
