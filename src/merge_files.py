"""
Merge Excel Files Module - Simple merge with NO aggregation.

Features:
- Upload 1 to 20 Excel files
- Merge all files (just stack rows)
- NO aggregation, NO grouping
- Color-coded rows by source file
- Download complete merged data as Excel/CSV
"""
import streamlit as st
import pandas as pd
import io
from typing import List, Tuple
from openpyxl.styles import PatternFill, Font


def read_excel_optimized(uploaded_file) -> pd.DataFrame:
    """Read Excel file with optimization for large files."""
    filename = uploaded_file.name.lower()
    
    if filename.endswith('.csv'):
        return pd.read_csv(uploaded_file, low_memory=False)
    else:
        return pd.read_excel(uploaded_file)


def generate_colored_excel(df: pd.DataFrame, file_sources: List[Tuple[str, int, int]]) -> bytes:
    """
    Generate Excel file with color-coded rows by source file.
    
    Args:
        df: The merged dataframe
        file_sources: List of (filename, start_row, end_row) tuples for color coding
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Merged Data')
        
        # Get the worksheet
        worksheet = writer.sheets['Merged Data']
        
        # Bold the header row
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        
        # Color code rows by source file
        # Define colors for different files (pastel colors)
        colors = [
            'E3F2FD',  # Light Blue
            'F3E5F5',  # Light Purple
            'E8F5E9',  # Light Green
            'FFF3E0',  # Light Orange
            'FCE4EC',  # Light Pink
            'F1F8E9',  # Light Lime
            'E0F2F1',  # Light Teal
            'FFF9C4',  # Light Yellow
            'FFEBEE',  # Light Red
            'E8EAF6',  # Light Indigo
        ]
        
        for idx, (filename, start_row, end_row) in enumerate(file_sources):
            color = colors[idx % len(colors)]
            fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            
            # Apply color to rows (add 2 because: 1 for header, 1 for 0-indexing)
            for row_idx in range(start_row + 2, end_row + 2):
                for cell in worksheet[row_idx]:
                    cell.fill = fill
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return output.getvalue()


def render_merge_files_page():
    """Render the Merge Excel Files page - SIMPLE MERGE ONLY."""
    st.title("📂 Merge Excel Files")
    st.markdown("""
    Upload **1 to 20 Excel files** and merge them into one file.
    - **NO aggregation** - just combines all rows from all files
    - All columns are preserved
    - Files are stacked vertically (row by row)
    - Download the complete merged data
    """)
    
    st.markdown("---")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Choose Excel/CSV files (1-20 files)",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        key="merge_file_uploader"
    )
    
    if uploaded_files:
        st.info(f"📁 **{len(uploaded_files)}** file(s) uploaded")
        
        # Show file list
        with st.expander("View uploaded files", expanded=False):
            for f in uploaded_files:
                size_kb = f.size / 1024
                if size_kb > 1024:
                    size_str = f"{size_kb/1024:.2f} MB"
                else:
                    size_str = f"{size_kb:.2f} KB"
                st.write(f"• **{f.name}** — {size_str}")
    
    st.markdown("---")
    
    # Process button
    process_btn = st.button("🚀 Merge Files", type="primary", use_container_width=True)
    
    if process_btn:
        if not uploaded_files:
            st.warning("⚠️ Please upload at least 1 file")
            return
        
        # UPDATED: Limit changed from 15 to 20
        if len(uploaded_files) > 20:
            st.warning("⚠️ Maximum 20 files allowed. Please remove some files.")
            return
        
        # Read all files
        all_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Reading {uploaded_file.name}...")
            progress_bar.progress((i + 1) / len(uploaded_files))
            
            try:
                df = read_excel_optimized(uploaded_file)
                
                if len(df) == 0:
                    st.warning(f"⚠️ **{uploaded_file.name}**: File is empty, skipping")
                    continue
                
                st.success(f"✅ **{uploaded_file.name}**: {len(df):,} rows, {len(df.columns)} columns")
                all_data.append((uploaded_file.name, df))
                
            except Exception as e:
                st.error(f"❌ **{uploaded_file.name}**: Error - {str(e)}")
        
        progress_bar.progress(100)
        status_text.text("Reading complete!")
        
        if not all_data:
            st.error("❌ No valid data found in any uploaded files.")
            return
        
        # Merge all data
        st.markdown("---")
        st.subheader("📊 Merging Files...")
        
        with st.spinner("Combining all files..."):
            # Track which rows came from which file for color coding
            file_sources = []
            current_row = 0
            
            # Simple concatenation - NO AGGREGATION
            dfs_to_merge = []
            for filename, df in all_data:
                dfs_to_merge.append(df)
                # Track file source for coloring
                file_sources.append((filename, current_row, current_row + len(df)))
                current_row += len(df)
            
            combined_df = pd.concat(dfs_to_merge, ignore_index=True, sort=False)
            
            total_input_rows = sum(len(df) for _, df in all_data)
            
            st.success(f"✅ Merged successfully!")
            st.info(f"**Result:** {len(combined_df):,} rows × {len(combined_df.columns)} columns")
            
            # Verify row count
            if len(combined_df) == total_input_rows:
                st.success(f"✅ All {total_input_rows:,} rows merged correctly!")
            else:
                st.warning(f"⚠️ Expected {total_input_rows:,} rows but got {len(combined_df):,}")
        
        # Store in session state
        st.session_state['merge_combined'] = combined_df
        st.session_state['merge_file_list'] = all_data
        st.session_state['merge_file_sources'] = file_sources
    
    # Show results if available
    if 'merge_combined' in st.session_state:
        combined_df = st.session_state['merge_combined']
        file_list = st.session_state.get('merge_file_list', [])
        file_sources = st.session_state.get('merge_file_sources', [])
        
        st.markdown("---")
        st.subheader("📋 Merged Data")
        
        # Stats
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", f"{len(combined_df):,}")
        with col2:
            st.metric("Total Columns", len(combined_df.columns))
        with col3:
            size_mb = combined_df.memory_usage(deep=True).sum() / (1024 * 1024)
            st.metric("Size", f"{size_mb:.1f} MB")
        
        # Show file breakdown
        with st.expander("📋 Files Breakdown", expanded=False):
            breakdown_data = []
            for filename, df in file_list:
                breakdown_data.append({
                    'File Name': filename,
                    'Rows': f"{len(df):,}",
                    'Columns': len(df.columns)
                })
            breakdown_df = pd.DataFrame(breakdown_data)
            st.dataframe(breakdown_df, use_container_width=True, hide_index=True)
        
        # Show color legend
        with st.expander("🎨 Color Legend (in downloaded Excel)", expanded=False):
            st.markdown("**Each file's rows will have a different background color in the Excel file:**")
            colors_display = [
                ('Light Blue', '#E3F2FD'),
                ('Light Purple', '#F3E5F5'),
                ('Light Green', '#E8F5E9'),
                ('Light Orange', '#FFF3E0'),
                ('Light Pink', '#FCE4EC'),
                ('Light Lime', '#F1F8E9'),
                ('Light Teal', '#E0F2F1'),
                ('Light Yellow', '#FFF9C4'),
                ('Light Red', '#FFEBEE'),
                ('Light Indigo', '#E8EAF6'),
            ]
            
            for idx, (filename, df) in enumerate(file_list):
                color_name, color_hex = colors_display[idx % len(colors_display)]
                st.markdown(f"**{idx + 1}. {filename}** → {color_name}")
        
        # Preview
        with st.expander("📋 Preview Merged Data (First 100 rows)", expanded=True):
            st.dataframe(combined_df.head(100), use_container_width=True)
        
        st.markdown("---")
        st.subheader("📥 Download Merged File")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Excel download with colors
            with st.spinner("Generating Excel file with color-coded rows..."):
                excel_bytes = generate_colored_excel(combined_df, file_sources)
            
            st.download_button(
                label=f"📊 Download Excel ({len(combined_df):,} rows, color-coded)",
                data=excel_bytes,
                file_name="merged_full_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
                help="Excel file with bold headers and color-coded rows by source file"
            )
        
        with col2:
            # CSV download
            csv_data = combined_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📄 Download CSV ({len(combined_df):,} rows)",
                data=csv_data,
                file_name="merged_full_data.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        # Clear button
        st.markdown("---")
        if st.button("🔄 Clear & Start Over", use_container_width=True):
            if 'merge_combined' in st.session_state:
                del st.session_state['merge_combined']
            if 'merge_file_list' in st.session_state:
                del st.session_state['merge_file_list']
            if 'merge_file_sources' in st.session_state:
                del st.session_state['merge_file_sources']
            st.rerun()
