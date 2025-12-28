import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
from pathlib import Path
import calendar

# Page configuration
st.set_page_config(
    page_title="Sales Target & Analysis App",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load custom CSS and JS
def load_css():
    try:
        with open('styles.css', 'r') as f:
            css = f.read()
        st.markdown(f'<style>{css}</style>', unsafe_allow_html=True)
        
        # Load JavaScript
        try:
            with open('script.js', 'r') as f:
                js = f.read()
            st.markdown(f'<script>{js}</script>', unsafe_allow_html=True)
        except FileNotFoundError:
            pass
        
            
    except FileNotFoundError:
        # If CSS file doesn't exist, use minimal inline styles
        st.markdown("""
        <style>
            .main { background-color: #ffffff; }
            @media (max-width: 768px) {
                .stColumns { flex-direction: column; }
            }
        </style>
        """, unsafe_allow_html=True)


# Initialize session state
if 'sales_data' not in st.session_state:
    st.session_state.sales_data = None
if 'ams_data' not in st.session_state:
    st.session_state.ams_data = None
if 'target_data' not in st.session_state:
    st.session_state.target_data = None

def parse_mmm_yr(date_str):
    """
    Parse various date formats to datetime.
    Accepts formats like:
    - 'Jan-2024', 'Feb-2024' (mmm-yr)
    - '2024-01-01', '2024/01/01' (ISO format)
    - '01/2024', '01-2024' (month/year)
    - 'January 2024', 'Jan 2024' (full/abbreviated month name)
    - '2024-01', '2024/01' (year-month)
    - Excel date serial numbers
    - Any other format pandas can parse
    """
    # If already datetime, normalize to first day of month
    if isinstance(date_str, pd.Timestamp) or isinstance(date_str, datetime):
        dt = pd.to_datetime(date_str)
        return pd.to_datetime(f"{dt.year}-{dt.month:02d}-01")
    
    # If it's already a datetime-like object, convert it
    if pd.isna(date_str):
        return pd.NaT
    
    # Convert to string for processing
    if not isinstance(date_str, str):
        date_str = str(date_str).strip()
    else:
        date_str = date_str.strip()
    
    # Try parsing "mmm-yr" format first (e.g., "Jan-2024", "Feb-2024")
    try:
        if '-' in date_str and len(date_str.split('-')) == 2:
            parts = date_str.split('-')
            month_part = parts[0].strip()
            year_part = parts[1].strip()
            
            # Check if month_part is a month abbreviation or name
            month_abbr_list = [m.lower() for m in calendar.month_abbr[1:]]
            month_name_list = [m.lower() for m in calendar.month_name[1:]]
            
            if month_part.lower() in month_abbr_list or month_part.lower() in month_name_list:
                if month_part.lower() in month_abbr_list:
                    month_num = month_abbr_list.index(month_part.lower()) + 1
                else:
                    month_num = month_name_list.index(month_part.lower()) + 1
                
                # Handle year part - could be 2-digit (yy) or 4-digit (yyyy)
                year = int(year_part)
                # If year is 2 digits (0-99), convert to 4 digits
                # Assume years 0-99 are 2000-2099
                if 0 <= year <= 99:
                    year = 2000 + year
                
                return pd.to_datetime(f"{year}-{month_num:02d}-01")
    except (ValueError, IndexError):
        pass
    
    # Try parsing "mm/yyyy" or "mm-yyyy" format (e.g., "01/2024", "12-2024")
    try:
        if ('/' in date_str or '-' in date_str) and len(date_str) <= 7:
            parts = date_str.replace('/', '-').split('-')
            if len(parts) == 2:
                first = parts[0].strip()
                second = parts[1].strip()
                
                # Check if first part is numeric and could be month (1-12)
                if first.isdigit() and second.isdigit():
                    month_val = int(first)
                    year_val = int(second)
                    
                    # Determine which is month and which is year
                    # If first < 13, likely month/year, else year/month
                    if 1 <= month_val <= 12 and len(second) == 4:
                        return pd.to_datetime(f"{year_val}-{month_val:02d}-01")
                    elif 1 <= year_val <= 12 and len(first) == 4:
                        return pd.to_datetime(f"{int(first)}-{year_val:02d}-01")
    except (ValueError, IndexError):
        pass
    
    # Try parsing "yyyy-mm" or "yyyy/mm" format (e.g., "2024-01", "2024/01")
    try:
        if ('-' in date_str or '/' in date_str) and len(date_str) == 7:
            parts = date_str.replace('/', '-').split('-')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                if len(parts[0]) == 4:  # Year first
                    return pd.to_datetime(f"{parts[0]}-{parts[1]}-01")
    except (ValueError, IndexError):
        pass
    
    # Try parsing with pandas (handles most standard formats)
    try:
        parsed = pd.to_datetime(date_str, errors='raise')
        # If parsed successfully, ensure it's set to first day of month
        if pd.notna(parsed):
            return pd.to_datetime(f"{parsed.year}-{parsed.month:02d}-01")
    except (ValueError, TypeError, OverflowError):
        pass
    
    # Last resort: try Excel date serial number
    try:
        if date_str.replace('.', '').isdigit():
            excel_date = float(date_str)
            # Excel epoch is 1899-12-30
            base_date = pd.to_datetime('1899-12-30')
            parsed = base_date + pd.Timedelta(days=excel_date)
            return pd.to_datetime(f"{parsed.year}-{parsed.month:02d}-01")
    except (ValueError, TypeError, OverflowError):
        pass
    
    # If all else fails, raise an error
    raise ValueError(f"Unable to parse date: {date_str}")

def format_to_mmm_yr(date_val):
    """Convert datetime to 'mmm-yr' format (e.g., 'Jan-2024')"""
    if pd.isna(date_val):
        return ''
    if isinstance(date_val, (pd.Timestamp, datetime)):
        return date_val.strftime('%b-%Y')
    # Try to parse it first
    try:
        parsed = pd.to_datetime(date_val)
        return parsed.strftime('%b-%Y')
    except:
        return str(date_val)

def calculate_ams(df, num_months=6, exclusion_threshold_percent=20):
    """
    Calculate Average Monthly Sales (AMS) for specified number of previous months,
    ignoring months with sales below the specified percentage threshold of AMS
    
    Parameters:
    - df: DataFrame with sales data
    - num_months: Number of previous months to use for AMS calculation (default: 6)
    - exclusion_threshold_percent: Percentage threshold below which months are excluded (default: 20)
    """
    # Parse Mth-yr to datetime (handles mmm-yr format)
    df = df.copy()
    df['Mth-yr'] = df['Mth-yr'].apply(parse_mmm_yr)
    
    # Get last N months
    max_date = df['Mth-yr'].max()
    months_ago = max_date - pd.DateOffset(months=num_months)
    last_months = df[df['Mth-yr'] >= months_ago].copy()
    
    # Group by Product, Customer Type, Township, Region
    grouped = last_months.groupby(['Product', 'Customer Type', 'Township', 'Region'])
    
    ams_results = []
    
    for (product, customer_type, township, region), group in grouped:
        # Sort by date
        group = group.sort_values('Mth-yr')
        
        # Calculate initial AMS (all months)
        monthly_sales = group.groupby('Mth-yr')['Sales Qty'].sum()
        initial_ams = monthly_sales.mean()
        
        # Filter out months with sales below the threshold percentage of initial AMS
        if initial_ams > 0:
            threshold = (exclusion_threshold_percent / 100) * initial_ams
            filtered_months = monthly_sales[monthly_sales >= threshold]
            
            # Recalculate AMS with filtered months
            if len(filtered_months) > 0:
                ams = filtered_months.mean()
                months_counted = len(filtered_months)
                months_excluded = len(monthly_sales) - months_counted
            else:
                # If all months are below threshold, use all months
                ams = initial_ams
                months_counted = len(monthly_sales)
                months_excluded = 0
        else:
            ams = 0
            months_counted = 0
            months_excluded = 0
        
        ams_results.append({
            'Product': product,
            'Customer Type': customer_type,
            'Township': township,
            'Region': region,
            'AMS': round(ams),  # Round to integer, no decimals
            'Months Counted': months_counted,
            'Months Excluded': months_excluded,
            'Total Months': len(monthly_sales)
        })
    
    df_result = pd.DataFrame(ams_results)
    # Add row number column starting from 1
    df_result.insert(0, 'No.', range(1, len(df_result) + 1))
    return df_result

def calculate_targets(ams_df, percentage_increase):
    """Calculate targets based on AMS and percentage increase"""
    target_df = ams_df.copy()
    target_df['Target Qty'] = (target_df['AMS'] * (1 + percentage_increase / 100)).round().astype(int)  # Round to integer, no decimals
    # Keep the row number column if it exists
    if 'No.' in target_df.columns:
        target_df = target_df[['No.', 'Product', 'Customer Type', 'Township', 'Region', 'Target Qty', 'AMS']]
    else:
        target_df = target_df[['Product', 'Customer Type', 'Township', 'Region', 'Target Qty', 'AMS']]
    return target_df

def create_template():
    """Create a sample template file with mmm-yr format"""
    # Generate last 12 months of sample data
    months = []
    products = ['Product A', 'Product B', 'Product C']
    customer_types = ['Retail', 'Wholesale', 'Corporate']
    townships = ['Township 1', 'Township 2', 'Township 3']
    regions = ['Region 1', 'Region 2']
    
    template_data = {
        'Mth-yr': [],
        'Product': [],
        'Customer Type': [],
        'Township': [],
        'Region': [],
        'Sales Qty': []
    }
    
    # Generate sample data for last 12 months
    today = datetime.now()
    for i in range(12):
        month_date = today - pd.DateOffset(months=11-i)
        month_str = month_date.strftime('%b-%Y')  # Format: Jan-2024
        
        # Create multiple rows per month with different combinations
        for product in products[:2]:  # Use first 2 products
            for customer_type in customer_types[:2]:  # Use first 2 customer types
                for township in townships[:2]:  # Use first 2 townships
                    for region in regions:
                        template_data['Mth-yr'].append(month_str)
                        template_data['Product'].append(product)
                        template_data['Customer Type'].append(customer_type)
                        template_data['Township'].append(township)
                        template_data['Region'].append(region)
                        # Random sales quantity between 50 and 500
                        template_data['Sales Qty'].append(np.random.randint(50, 500))
    
    return pd.DataFrame(template_data)

def main():
    # Load CSS
    load_css()
    
    # Custom top navigation banner
    st.markdown("""
    <div id="custom-top-banner">
        <div class="banner-content">
            <div class="banner-tabs">
                <button class="banner-tab active" onclick="switchTab(0)">Target Planning & AMS</button>
                <button class="banner-tab" onclick="switchTab(1)">Sales Analysis</button>
            </div>
            <button id="toggle-streamlit-ui" class="toggle-ui-btn" onclick="toggleStreamlitUI()" title="Toggle Streamlit UI">
                <span>⚙</span>
            </button>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Top navigation tabs (hidden but functional)
    tab1, tab2 = st.tabs(["Target Planning & AMS", "Sales Analysis"])
    
    with tab1:
        target_planning_page()
    
    with tab2:
        sales_analysis_page()

def target_planning_page():
    st.header("🎯 Target Planning & AMS")
    
    # Template download section
    st.subheader("1. Download Template")
    template_df = create_template()
    
    # Convert to Excel for download with proper formatting
    template_buffer = io.BytesIO()
    with pd.ExcelWriter(template_buffer, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Sales Data')
        worksheet = writer.sheets['Sales Data']
        # Set column widths
        worksheet.set_column('A:A', 15)  # Mth-yr column
        worksheet.set_column('B:B', 20)  # Product
        worksheet.set_column('C:C', 15)  # Customer Type
        worksheet.set_column('D:D', 15)  # Township
        worksheet.set_column('E:E', 15)  # Region
        worksheet.set_column('F:F', 12)  # Sales Qty
        # Format header row
        header_format = writer.book.add_format({
            'bold': True,
            'bg_color': '#366092',
            'font_color': 'white'
        })
        for col_num, value in enumerate(template_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
    template_buffer.seek(0)
    
    st.download_button(
        label="📥 Download Sales Data Template",
        data=template_buffer,
        file_name="sales_data_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.info("📋 Template contains: Mth-yr (any date format with month and year is accepted, e.g., Jan-2024, 2024-01-01, 01/2024, January 2024), Product, Customer Type, Township, Region, Sales Qty")
    
    # File upload section
    st.subheader("2. Upload Sales Data")
    uploaded_file = st.file_uploader(
        "Choose a file",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your sales data file"
    )
    
    if uploaded_file is not None:
        try:
            # Initialize progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 1: File upload detected (0-20%)
            status_text.text("📤 File uploaded, starting processing...")
            progress_bar.progress(10)
            
            # Step 2: Read the file (20-50%)
            status_text.text("📖 Reading file...")
            progress_bar.progress(20)
            
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            progress_bar.progress(50)
            status_text.text("✅ File read successfully, validating data...")
            
            # Step 3: Validate columns (50-60%)
            required_columns = ['Mth-yr', 'Product', 'Customer Type', 'Township', 'Region', 'Sales Qty']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            progress_bar.progress(60)
            
            if missing_columns:
                progress_bar.empty()
                status_text.empty()
                st.error(f"Missing required columns: {', '.join(missing_columns)}")
            else:
                # Step 4: Validate and parse dates (60-90%)
                status_text.text("📅 Parsing and validating dates...")
                progress_bar.progress(60)
                
                date_errors = []
                df_copy = df.copy()
                
                total_rows = len(df_copy)
                # Try to parse dates and collect errors
                for idx, date_val in enumerate(df_copy['Mth-yr']):
                    try:
                        parse_mmm_yr(date_val)
                    except (ValueError, TypeError) as e:
                        date_errors.append({
                            'row': idx + 2,  # +2 because of header and 0-based index
                            'value': date_val,
                            'error': str(e)
                        })
                    
                    # Update progress during date parsing
                    if (idx + 1) % max(1, total_rows // 10) == 0 or idx == total_rows - 1:
                        progress = 60 + int((idx + 1) / total_rows * 30)
                        progress_bar.progress(progress)
                
                progress_bar.progress(90)
                status_text.text("🔍 Finalizing data processing...")
                
                if date_errors:
                    st.warning(f"⚠️ Found {len(date_errors)} date(s) that couldn't be parsed. Please check the dates in your file.")
                    with st.expander("View date parsing errors"):
                        error_df = pd.DataFrame(date_errors)
                        st.dataframe(error_df, use_container_width=True)
                
                # Step 5: Complete (90-100%)
                st.session_state.sales_data = df_copy
                progress_bar.progress(100)
                status_text.text("✅ Data loaded successfully!")
                
                # Clear progress bar and status after a brief moment
                import time
                time.sleep(0.3)
                progress_bar.empty()
                status_text.empty()
                
                st.success(f"✅ Data loaded successfully! {len(df)} rows")
                
                # Display data preview
                with st.expander("Preview Data"):
                    # Show original dates in preview
                    preview_df = df_copy.head(10).copy()
                    st.dataframe(preview_df, use_container_width=True)
                
                # Filters section
                st.subheader("3. Apply Filters")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    products = st.multiselect(
                        "Product",
                        options=sorted(df['Product'].unique()),
                        default=sorted(df['Product'].unique())
                    )
                
                with col2:
                    customer_types = st.multiselect(
                        "Customer Type",
                        options=sorted(df['Customer Type'].unique()),
                        default=sorted(df['Customer Type'].unique())
                    )
                
                with col4:
                    regions = st.multiselect(
                        "Region",
                        options=sorted(df['Region'].unique()),
                        default=sorted(df['Region'].unique())
                    )
                
                with col3:
                    # Smart township options: If regions are selected,
                    # show townships from those regions and auto-select if empty
                    if len(regions) > 0:
                        # Regions are selected, so show only townships from selected regions
                        available_townships = sorted(df[df['Region'].isin(regions)]['Township'].unique())
                        townships = st.multiselect(
                            "Township",
                            options=available_townships,
                            default=available_townships,  # Default to all townships in selected regions
                            help="Select specific townships, or leave all selected to include all townships in the chosen region(s)"
                        )
                        # If user deselects all townships, automatically include all from selected regions
                        if len(townships) == 0:
                            townships = available_townships
                            st.info(f"ℹ️ Automatically including all {len(townships)} township(s) from selected region(s): {', '.join(regions)}")
                    else:
                        # No regions selected, show all townships
                        townships = st.multiselect(
                            "Township",
                            options=sorted(df['Township'].unique()),
                            default=sorted(df['Township'].unique())
                        )
                
                # Filter the data with progress bar
                filter_progress = st.progress(0)
                filter_status = st.empty()
                
                filter_status.text("🔄 Applying filters...")
                filter_progress.progress(30)
                
                filtered_df = df[
                    (df['Product'].isin(products)) &
                    (df['Customer Type'].isin(customer_types)) &
                    (df['Township'].isin(townships)) &
                    (df['Region'].isin(regions))
                ]
                
                filter_progress.progress(100)
                filter_status.text("✅ Filters applied!")
                
                import time
                time.sleep(0.2)
                filter_progress.empty()
                filter_status.empty()
                
                if len(filtered_df) == 0:
                    st.warning("No data matches the selected filters.")
                else:
                    # Calculate AMS
                    st.subheader("4. Calculate AMS & Targets")
                    
                    # AMS Calculation Parameters
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        num_months = st.number_input(
                            "Number of Previous Months for AMS Calculation",
                            min_value=1,
                            max_value=24,
                            value=6,
                            step=1,
                            help="Select how many previous months of sales data to use for AMS calculation"
                        )
                    
                    with col2:
                        exclusion_threshold = st.number_input(
                            "Exclusion Threshold (%)",
                            min_value=0.0,
                            max_value=100.0,
                            value=20.0,
                            step=1.0,
                            help="Months with sales below this percentage of AMS will be excluded from calculation (to reflect actual demand)"
                        )
                    
                    st.markdown("---")
                    
                    percentage_increase = st.number_input(
                        "Target % Increase on AMS",
                        min_value=0.0,
                        max_value=1000.0,
                        value=10.0,
                        step=0.1,
                        help="Enter the percentage increase for target calculation"
                    )
                    
                    if st.button("Calculate AMS & Targets", type="primary"):
                        # Initialize progress bar
                        calc_progress = st.progress(0)
                        calc_status = st.empty()
                        
                        # Step 1: Calculate AMS
                        calc_status.text("📊 Calculating Average Monthly Sales (AMS)...")
                        calc_progress.progress(20)
                        ams_df = calculate_ams(filtered_df, num_months=num_months, exclusion_threshold_percent=exclusion_threshold)
                        st.session_state.ams_data = ams_df
                        calc_progress.progress(60)
                        
                        # Step 2: Calculate Targets
                        calc_status.text("🎯 Calculating target quantities...")
                        calc_progress.progress(70)
                        target_df = calculate_targets(ams_df, percentage_increase)
                        st.session_state.target_data = target_df
                        calc_progress.progress(90)
                        
                        # Step 3: Complete
                        calc_status.text("✅ Calculation completed!")
                        calc_progress.progress(100)
                        
                        import time
                        time.sleep(0.3)
                        calc_progress.empty()
                        calc_status.empty()
                        
                        st.success(f"✅ Calculation completed! Used {num_months} months with {exclusion_threshold}% exclusion threshold.")
                    
                    # Display results
                    if st.session_state.ams_data is not None:
                        st.subheader("5. Results")
                        
                        tab1, tab2 = st.tabs(["AMS Data", "Target Data"])
                        
                        with tab1:
                            # View selection
                            ams_view_option = st.radio(
                                "View Options",
                                ["Detailed View", "Filtered View"],
                                horizontal=True,
                                key="ams_view_option"
                            )
                            
                            if ams_view_option == "Detailed View":
                                # Format AMS data for display (no decimals)
                                ams_display = st.session_state.ams_data.copy()
                                if 'AMS' in ams_display.columns:
                                    ams_display['AMS'] = ams_display['AMS'].astype(int)
                                st.dataframe(ams_display, use_container_width=True)
                                
                                # Download AMS
                                ams_buffer = io.BytesIO()
                                with pd.ExcelWriter(ams_buffer, engine='xlsxwriter') as writer:
                                    st.session_state.ams_data.to_excel(writer, index=False, sheet_name='AMS')
                                    # Format the worksheet
                                    worksheet = writer.sheets['AMS']
                                    # Format AMS column as integer (no decimals)
                                    if 'AMS' in st.session_state.ams_data.columns:
                                        ams_col = list(st.session_state.ams_data.columns).index('AMS') + 1
                                        num_format = writer.book.add_format({'num_format': '0'})
                                        worksheet.set_column(ams_col, ams_col, None, num_format)
                                ams_buffer.seek(0)
                                
                                st.download_button(
                                    label="📥 Download AMS Data",
                                    data=ams_buffer,
                                    file_name=f"AMS_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            
                            else:  # Filtered View
                                # Filters for Filtered View
                                st.markdown("### Filters")
                                col1, col2, col3 = st.columns(3)
                                
                                with col1:
                                    ams_summary_products = st.multiselect(
                                        "Product",
                                        options=sorted(st.session_state.ams_data['Product'].unique()),
                                        default=sorted(st.session_state.ams_data['Product'].unique()),
                                        key="ams_summary_products"
                                    )
                                
                                with col2:
                                    ams_summary_customer_types = st.multiselect(
                                        "Customer Type",
                                        options=sorted(st.session_state.ams_data['Customer Type'].unique()),
                                        default=sorted(st.session_state.ams_data['Customer Type'].unique()),
                                        key="ams_summary_customer_types"
                                    )
                                
                                with col3:
                                    ams_summary_regions = st.multiselect(
                                        "Region",
                                        options=sorted(st.session_state.ams_data['Region'].unique()),
                                        default=sorted(st.session_state.ams_data['Region'].unique()),
                                        key="ams_summary_regions"
                                    )
                                
                                # Filter AMS data by Product, Customer Type, and Region with progress
                                summary_progress = st.progress(0)
                                summary_status = st.empty()
                                
                                summary_status.text("🔄 Processing filters...")
                                summary_progress.progress(20)
                                
                                filtered_ams = st.session_state.ams_data[
                                    (st.session_state.ams_data['Product'].isin(ams_summary_products)) &
                                    (st.session_state.ams_data['Customer Type'].isin(ams_summary_customer_types)) &
                                    (st.session_state.ams_data['Region'].isin(ams_summary_regions))
                                ]
                                
                                summary_progress.progress(40)
                                summary_status.text("📊 Aggregating data...")
                                
                                if len(filtered_ams) == 0:
                                    summary_progress.empty()
                                    summary_status.empty()
                                    st.warning("No data matches the selected filters.")
                                else:
                                    # 1. Product Summary (aggregate by Product)
                                    ams_product_summary = filtered_ams.groupby('Product').agg({
                                        'AMS': 'sum'
                                    }).reset_index()
                                    
                                    ams_product_summary['AMS'] = ams_product_summary['AMS'].astype(int)
                                    ams_product_summary = ams_product_summary.sort_values('AMS', ascending=False).reset_index(drop=True)
                                    ams_product_summary.insert(0, 'No.', range(1, len(ams_product_summary) + 1))
                                    ams_product_summary = ams_product_summary[['No.', 'Product', 'AMS']]
                                    
                                    summary_progress.progress(60)
                                    summary_status.text("📊 Creating Product Summary by Region by Customer Type...")
                                    
                                    # 2. Product Summary by Region by Customer Type (for selected regions and customer types)
                                    ams_by_region_customer = filtered_ams.groupby(['Product', 'Region', 'Customer Type']).agg({
                                        'AMS': 'sum'
                                    }).reset_index()
                                    
                                    ams_by_region_customer['AMS'] = ams_by_region_customer['AMS'].astype(int)
                                    ams_by_region_customer = ams_by_region_customer.sort_values('AMS', ascending=False).reset_index(drop=True)
                                    ams_by_region_customer.insert(0, 'No.', range(1, len(ams_by_region_customer) + 1))
                                    ams_by_region_customer = ams_by_region_customer[['No.', 'Product', 'Region', 'Customer Type', 'AMS']]
                                    
                                    summary_progress.progress(100)
                                    summary_status.text("✅ Processing complete!")
                                    
                                    import time
                                    time.sleep(0.2)
                                    summary_progress.empty()
                                    summary_status.empty()
                                    
                                    # Display Product Summary
                                    st.markdown("### Product Summary (Total AMS by Product)")
                                    filter_info = []
                                    if len(ams_summary_products) < len(st.session_state.ams_data['Product'].unique()):
                                        filter_info.append(f"Products: {', '.join(ams_summary_products)}")
                                    if len(ams_summary_customer_types) < len(st.session_state.ams_data['Customer Type'].unique()):
                                        filter_info.append(f"Customer Type(s): {', '.join(ams_summary_customer_types)}")
                                    if len(ams_summary_regions) < len(st.session_state.ams_data['Region'].unique()):
                                        filter_info.append(f"Region(s): {', '.join(ams_summary_regions)}")
                                    if filter_info:
                                        st.info(f"Showing AMS totals for: {' | '.join(filter_info)}")
                                    st.dataframe(ams_product_summary, use_container_width=True)
                                    
                                    # Display Product Summary by Region by Customer Type
                                    st.markdown("### Product Summary by Region by Customer Type")
                                    st.dataframe(ams_by_region_customer, use_container_width=True)
                                    
                                    # Download all summaries
                                    ams_summary_buffer = io.BytesIO()
                                    with pd.ExcelWriter(ams_summary_buffer, engine='xlsxwriter') as writer:
                                        ams_product_summary.to_excel(writer, index=False, sheet_name='AMS Product Summary')
                                        ams_by_region_customer.to_excel(writer, index=False, sheet_name='AMS by Region by Customer Type')
                                        
                                        # Format all worksheets
                                        num_format = writer.book.add_format({'num_format': '0'})
                                        
                                        # Format AMS Product Summary
                                        worksheet1 = writer.sheets['AMS Product Summary']
                                        ams_col1 = list(ams_product_summary.columns).index('AMS') + 1
                                        worksheet1.set_column(ams_col1, ams_col1, None, num_format)
                                        
                                        # Format AMS by Region by Customer Type
                                        worksheet2 = writer.sheets['AMS by Region by Customer Type']
                                        ams_col2 = list(ams_by_region_customer.columns).index('AMS') + 1
                                        worksheet2.set_column(ams_col2, ams_col2, None, num_format)
                                    ams_summary_buffer.seek(0)
                                    
                                    st.download_button(
                                        label="📥 Download All AMS Summaries",
                                        data=ams_summary_buffer,
                                        file_name=f"AMS_Summaries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with tab2:
                            # View selection
                            view_option = st.radio(
                                "View Options",
                                ["Detailed View", "Filtered View"],
                                horizontal=True,
                                key="target_view_option"
                            )
                            
                            if view_option == "Detailed View":
                                # Format Target data for display (no decimals)
                                target_display = st.session_state.target_data.copy()
                                if 'Target Qty' in target_display.columns:
                                    target_display['Target Qty'] = target_display['Target Qty'].astype(int)
                                if 'AMS' in target_display.columns:
                                    target_display['AMS'] = target_display['AMS'].astype(int)
                                st.dataframe(target_display, use_container_width=True)
                                
                                # Download Targets
                                target_buffer = io.BytesIO()
                                with pd.ExcelWriter(target_buffer, engine='xlsxwriter') as writer:
                                    st.session_state.target_data.to_excel(writer, index=False, sheet_name='Targets')
                                    # Format the worksheet
                                    worksheet = writer.sheets['Targets']
                                    # Format Target Qty and AMS columns as integer (no decimals)
                                    num_format = writer.book.add_format({'num_format': '0'})
                                    if 'Target Qty' in st.session_state.target_data.columns:
                                        target_col = list(st.session_state.target_data.columns).index('Target Qty') + 1
                                        worksheet.set_column(target_col, target_col, None, num_format)
                                    if 'AMS' in st.session_state.target_data.columns:
                                        ams_col = list(st.session_state.target_data.columns).index('AMS') + 1
                                        worksheet.set_column(ams_col, ams_col, None, num_format)
                                target_buffer.seek(0)
                                
                                st.download_button(
                                    label="📥 Download Target Data",
                                    data=target_buffer,
                                    file_name=f"Target_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            
                            else:  # Filtered View
                                # Filters for Filtered View
                                st.markdown("### Filters")
                                col1, col2, col3 = st.columns(3)
                                
                                with col1:
                                    summary_products = st.multiselect(
                                        "Product",
                                        options=sorted(st.session_state.target_data['Product'].unique()),
                                        default=sorted(st.session_state.target_data['Product'].unique()),
                                        key="summary_products"
                                    )
                                
                                with col2:
                                    summary_customer_types = st.multiselect(
                                        "Customer Type",
                                        options=sorted(st.session_state.target_data['Customer Type'].unique()),
                                        default=sorted(st.session_state.target_data['Customer Type'].unique()),
                                        key="summary_customer_types"
                                    )
                                
                                with col3:
                                    summary_regions = st.multiselect(
                                        "Region",
                                        options=sorted(st.session_state.target_data['Region'].unique()),
                                        default=sorted(st.session_state.target_data['Region'].unique()),
                                        key="summary_regions"
                                    )
                                
                                # Filter target data by Product, Customer Type, and Region with progress
                                target_summary_progress = st.progress(0)
                                target_summary_status = st.empty()
                                
                                target_summary_status.text("🔄 Processing filters...")
                                target_summary_progress.progress(20)
                                
                                filtered_target = st.session_state.target_data[
                                    (st.session_state.target_data['Product'].isin(summary_products)) &
                                    (st.session_state.target_data['Customer Type'].isin(summary_customer_types)) &
                                    (st.session_state.target_data['Region'].isin(summary_regions))
                                ]
                                
                                target_summary_progress.progress(40)
                                target_summary_status.text("📊 Aggregating data...")
                                
                                if len(filtered_target) == 0:
                                    target_summary_progress.empty()
                                    target_summary_status.empty()
                                    st.warning("No data matches the selected filters.")
                                else:
                                    # 1. Product Summary (aggregate by Product)
                                    product_summary = filtered_target.groupby('Product').agg({
                                        'Target Qty': 'sum',
                                        'AMS': 'sum'
                                    }).reset_index()
                                    
                                    product_summary['Target Qty'] = product_summary['Target Qty'].astype(int)
                                    product_summary['AMS'] = product_summary['AMS'].astype(int)
                                    product_summary = product_summary.sort_values('Target Qty', ascending=False).reset_index(drop=True)
                                    product_summary.insert(0, 'No.', range(1, len(product_summary) + 1))
                                    product_summary = product_summary[['No.', 'Product', 'Target Qty', 'AMS']]
                                    
                                    target_summary_progress.progress(60)
                                    target_summary_status.text("📊 Creating Product Summary by Region by Customer Type...")
                                    
                                    # 2. Product Summary by Region by Customer Type (for selected regions and customer types)
                                    product_by_region_customer = filtered_target.groupby(['Product', 'Region', 'Customer Type']).agg({
                                        'Target Qty': 'sum',
                                        'AMS': 'sum'
                                    }).reset_index()
                                    
                                    product_by_region_customer['Target Qty'] = product_by_region_customer['Target Qty'].astype(int)
                                    product_by_region_customer['AMS'] = product_by_region_customer['AMS'].astype(int)
                                    product_by_region_customer = product_by_region_customer.sort_values('Target Qty', ascending=False).reset_index(drop=True)
                                    product_by_region_customer.insert(0, 'No.', range(1, len(product_by_region_customer) + 1))
                                    product_by_region_customer = product_by_region_customer[['No.', 'Product', 'Region', 'Customer Type', 'Target Qty', 'AMS']]
                                    
                                    target_summary_progress.progress(100)
                                    target_summary_status.text("✅ Processing complete!")
                                    
                                    import time
                                    time.sleep(0.2)
                                    target_summary_progress.empty()
                                    target_summary_status.empty()
                                    
                                    # Display Product Summary
                                    st.markdown("### Product Summary (Total Target Qty & AMS by Product)")
                                    filter_info = []
                                    if len(summary_products) < len(st.session_state.target_data['Product'].unique()):
                                        filter_info.append(f"Products: {', '.join(summary_products)}")
                                    if len(summary_customer_types) < len(st.session_state.target_data['Customer Type'].unique()):
                                        filter_info.append(f"Customer Type(s): {', '.join(summary_customer_types)}")
                                    if len(summary_regions) < len(st.session_state.target_data['Region'].unique()):
                                        filter_info.append(f"Region(s): {', '.join(summary_regions)}")
                                    if filter_info:
                                        st.info(f"Showing totals for: {' | '.join(filter_info)}")
                                    st.dataframe(product_summary, use_container_width=True)
                                    
                                    # Display Product Summary by Region by Customer Type
                                    st.markdown("### Product Summary by Region by Customer Type")
                                    st.dataframe(product_by_region_customer, use_container_width=True)
                                    
                                    # Download all summaries
                                    summary_buffer = io.BytesIO()
                                    with pd.ExcelWriter(summary_buffer, engine='xlsxwriter') as writer:
                                        product_summary.to_excel(writer, index=False, sheet_name='Product Summary')
                                        product_by_region_customer.to_excel(writer, index=False, sheet_name='Product by Region by Customer Type')
                                        
                                        # Format all worksheets
                                        num_format = writer.book.add_format({'num_format': '0'})
                                        
                                        # Format Product Summary
                                        worksheet1 = writer.sheets['Product Summary']
                                        target_col = list(product_summary.columns).index('Target Qty') + 1
                                        ams_col = list(product_summary.columns).index('AMS') + 1
                                        worksheet1.set_column(target_col, target_col, None, num_format)
                                        worksheet1.set_column(ams_col, ams_col, None, num_format)
                                        
                                        # Format Product by Region by Customer Type
                                        worksheet2 = writer.sheets['Product by Region by Customer Type']
                                        target_col2 = list(product_by_region_customer.columns).index('Target Qty') + 1
                                        ams_col2 = list(product_by_region_customer.columns).index('AMS') + 1
                                        worksheet2.set_column(target_col2, target_col2, None, num_format)
                                        worksheet2.set_column(ams_col2, ams_col2, None, num_format)
                                    
                                    summary_buffer.seek(0)
                                    
                                    st.download_button(
                                        label="📥 Download All Summaries",
                                        data=summary_buffer,
                                        file_name=f"Target_Summaries_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
        
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")

def get_filter_text(products_filter, customer_types_filter, townships_filter, regions_filter, df):
    """Generate filter text for chart annotations"""
    filter_parts = []
    all_products = sorted(df['Product'].unique())
    all_customer_types = sorted(df['Customer Type'].unique())
    all_townships = sorted(df['Township'].unique())
    all_regions = sorted(df['Region'].unique())
    
    if len(products_filter) < len(all_products):
        prod_text = ', '.join(products_filter[:3])
        if len(products_filter) > 3:
            prod_text += f' (+{len(products_filter)-3} more)'
        filter_parts.append(f"Products: {prod_text}")
    if len(customer_types_filter) < len(all_customer_types):
        filter_parts.append(f"Customer Types: {', '.join(customer_types_filter)}")
    if len(townships_filter) < len(all_townships):
        township_text = ', '.join(townships_filter[:3])
        if len(townships_filter) > 3:
            township_text += f' (+{len(townships_filter)-3} more)'
        filter_parts.append(f"Townships: {township_text}")
    if len(regions_filter) < len(all_regions):
        filter_parts.append(f"Regions: {', '.join(regions_filter)}")
    
    if filter_parts:
        return " | ".join(filter_parts)
    return "All data (no filters applied)"

def add_filter_annotation(fig, filter_text):
    """Add filter information as annotation to chart"""
    fig.add_annotation(
        text=f"Filters: {filter_text}",
        xref="paper", yref="paper",
        x=0.5, y=-0.12,
        xanchor="center", yanchor="top",
        showarrow=False,
        font=dict(size=9, color="#808495"),
        bgcolor="rgba(255,255,255,0.9)",
        bordercolor="#dadce0",
        borderwidth=1,
        borderpad=4
    )
    # Increase bottom margin to accommodate annotation
    fig.update_layout(margin=dict(b=60))
    return fig

def sales_analysis_page():
    st.header("📈 Sales Analysis")
    
    # Check if data is loaded
    if st.session_state.sales_data is None:
        st.warning("⚠️ Please upload sales data in the Target Planning section first.")
        return
    
    df = st.session_state.sales_data.copy()
    # Parse Mth-yr to datetime (handles mmm-yr format)
    df['Mth-yr'] = df['Mth-yr'].apply(parse_mmm_yr)
    
    # Filters
    st.subheader("Filters")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        products_filter = st.multiselect(
            "Product",
            options=sorted(df['Product'].unique()),
            default=sorted(df['Product'].unique()),
            key="analysis_products"
        )
    
    with col2:
        customer_types_filter = st.multiselect(
            "Customer Type",
            options=sorted(df['Customer Type'].unique()),
            default=sorted(df['Customer Type'].unique()),
            key="analysis_customer_types"
        )
    
    with col4:
        regions_filter = st.multiselect(
            "Region",
            options=sorted(df['Region'].unique()),
            default=sorted(df['Region'].unique()),
            key="analysis_regions"
        )
    
    with col3:
        # Smart township options: If regions are selected,
        # show townships from those regions and auto-select if empty
        if len(regions_filter) > 0:
            # Regions are selected, so show only townships from selected regions
            available_townships = sorted(df[df['Region'].isin(regions_filter)]['Township'].unique())
            townships_filter = st.multiselect(
                "Township",
                options=available_townships,
                default=available_townships,  # Default to all townships in selected regions
                help="Select specific townships, or leave all selected to include all townships in the chosen region(s)",
                key="analysis_townships"
            )
            # If user deselects all townships, automatically include all from selected regions
            if len(townships_filter) == 0:
                townships_filter = available_townships
                st.info(f"ℹ️ Automatically including all {len(townships_filter)} township(s) from selected region(s): {', '.join(regions_filter)}")
        else:
            # No regions selected, show all townships
            townships_filter = st.multiselect(
                "Township",
                options=sorted(df['Township'].unique()),
                default=sorted(df['Township'].unique()),
                key="analysis_townships"
            )
    
    # Filter data with progress bar
    analysis_progress = st.progress(0)
    analysis_status = st.empty()
    
    analysis_status.text("🔄 Applying filters...")
    analysis_progress.progress(30)
    
    filtered_df = df[
        (df['Product'].isin(products_filter)) &
        (df['Customer Type'].isin(customer_types_filter)) &
        (df['Township'].isin(townships_filter)) &
        (df['Region'].isin(regions_filter))
    ]
    
    analysis_progress.progress(70)
    analysis_status.text("📊 Preparing analysis data...")
    
    # Get filter text for charts
    filter_text = get_filter_text(products_filter, customer_types_filter, townships_filter, regions_filter, df)
    
    analysis_progress.progress(100)
    analysis_status.text("✅ Filters applied successfully!")
    
    import time
    time.sleep(0.2)
    analysis_progress.empty()
    analysis_status.empty()
    
    if len(filtered_df) == 0:
        st.warning("No data matches the selected filters.")
        return
    
    # Get date range from filtered data for date pickers
    min_date = filtered_df['Mth-yr'].min().date()
    max_date = filtered_df['Mth-yr'].max().date()
    
    # Ensure we have valid dates and a reasonable range for the date picker
    # Streamlit date_input needs a wider range to show year dropdown properly
    picker_min = datetime(2000, 1, 1).date()
    picker_max = datetime(2100, 12, 31).date()
    
    # Use actual data range but ensure it's within picker range
    actual_min = max(min_date, picker_min) if min_date >= picker_min else min_date
    actual_max = min(max_date, picker_max) if max_date <= picker_max else max_date
    
    # Analysis tabs
    tab1, tab2, tab3, tab4 = st.tabs(["Period Comparison", "Region Comparison", "Township Comparison", "Product Analysis"])
    
    with tab1:
        st.subheader("Period Comparison")
        
        # Period 1 and Period 2 selection for Period Comparison
        st.markdown("### Time Period Selection")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Period 1**")
            period1_start = st.date_input(
                "Period 1 Start",
                value=actual_min,
                min_value=picker_min,
                max_value=picker_max,
                key="period1_start"
            )
            period1_end = st.date_input(
                "Period 1 End",
                value=actual_max,
                min_value=picker_min,
                max_value=picker_max,
                key="period1_end"
            )
        
        with col2:
            st.markdown("**Period 2**")
            period2_start = st.date_input(
                "Period 2 Start",
                value=actual_min,
                min_value=picker_min,
                max_value=picker_max,
                key="period2_start"
            )
            period2_end = st.date_input(
                "Period 2 End",
                value=actual_max,
                min_value=picker_min,
                max_value=picker_max,
                key="period2_end"
            )
        
        # Validate that selected dates are within actual data range
        if period1_start < min_date:
            period1_start = min_date
        if period1_end > max_date:
            period1_end = max_date
        if period2_start < min_date:
            period2_start = min_date
        if period2_end > max_date:
            period2_end = max_date
        
        # Convert to datetime for filtering
        period1_start_dt = pd.to_datetime(period1_start)
        period1_end_dt = pd.to_datetime(period1_end)
        period2_start_dt = pd.to_datetime(period2_start)
        period2_end_dt = pd.to_datetime(period2_end)
        
        period1_data = filtered_df[
            (filtered_df['Mth-yr'] >= period1_start_dt) &
            (filtered_df['Mth-yr'] <= period1_end_dt)
        ]
        period2_data = filtered_df[
            (filtered_df['Mth-yr'] >= period2_start_dt) &
            (filtered_df['Mth-yr'] <= period2_end_dt)
        ]
        
        # Process period comparison with progress
        tab2_progress = st.progress(0)
        tab2_status = st.empty()
        tab2_status.text("📊 Processing period comparison...")
        tab2_progress.progress(50)
        
        if len(period1_data) > 0 and len(period2_data) > 0:
            period1_total = period1_data['Sales Qty'].sum()
            period2_total = period2_data['Sales Qty'].sum()
            change = period2_total - period1_total
            change_pct = (change / period1_total * 100) if period1_total > 0 else 0
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Period 1 Total", f"{period1_total:,.0f}")
            with col2:
                st.metric("Period 2 Total", f"{period2_total:,.0f}")
            with col3:
                st.metric("Change", f"{change:,.0f}", f"{change_pct:.1f}%")
            
            # Comparison chart
            comparison_data = pd.DataFrame({
                'Period': ['Period 1', 'Period 2'],
                'Sales Qty': [period1_total, period2_total]
            })
            
            # Format time period for title
            period1_text = f"Period 1: {period1_start.strftime('%b %d, %Y')} to {period1_end.strftime('%b %d, %Y')}"
            period2_text = f"Period 2: {period2_start.strftime('%b %d, %Y')} to {period2_end.strftime('%b %d, %Y')}"
            time_period_text = f"{period1_text} | {period2_text}"
            
            fig_comparison = px.bar(
                comparison_data,
                x='Period',
                y='Sales Qty',
                title='Period Comparison',
                text='Sales Qty'
            )
            fig_comparison.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            # Add time period to title
            fig_comparison.update_layout(
                height=400,
                title_text=f"Period Comparison<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                title_x=0.5
            )
            st.plotly_chart(fig_comparison, use_container_width=True)
            
            # Monthly breakdown
            period1_monthly = period1_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
            period1_monthly['Period'] = 'Period 1'
            period2_monthly = period2_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
            period2_monthly['Period'] = 'Period 2'
            
            tab2_progress.progress(80)
            tab2_status.text("📊 Generating comparison charts...")
            
            monthly_comparison = pd.concat([period1_monthly, period2_monthly])
            
            tab2_progress.progress(100)
            tab2_status.text("✅ Data processed!")
            import time
            time.sleep(0.1)
            tab2_progress.empty()
            tab2_status.empty()
            
            fig_monthly = px.line(
                monthly_comparison,
                x='Mth-yr',
                y='Sales Qty',
                color='Period',
                title='Monthly Breakdown by Period',
                markers=True
            )
            # Add time period to title
            fig_monthly.update_layout(
                height=400,
                title_text=f"Monthly Breakdown by Period<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                title_x=0.5
            )
            st.plotly_chart(fig_monthly, use_container_width=True)
            
            # Comparison table
            period1_avg = int(period1_data.groupby('Mth-yr')['Sales Qty'].sum().mean())
            period2_avg = int(period2_data.groupby('Mth-yr')['Sales Qty'].sum().mean())
            avg_change = period2_avg - period1_avg
            
            comparison_table = pd.DataFrame({
                'Metric': ['Total Sales', 'Average Monthly Sales', 'Number of Months'],
                'Period 1': [
                    int(period1_total),
                    period1_avg,
                    len(period1_data['Mth-yr'].unique())
                ],
                'Period 2': [
                    int(period2_total),
                    period2_avg,
                    len(period2_data['Mth-yr'].unique())
                ],
                'Change': [
                    int(change),
                    avg_change,
                    len(period2_data['Mth-yr'].unique()) - len(period1_data['Mth-yr'].unique())
                ]
            })
            
            st.dataframe(comparison_table, use_container_width=True)
            
            # Download comparison
            comparison_buffer = io.BytesIO()
            with pd.ExcelWriter(comparison_buffer, engine='xlsxwriter') as writer:
                comparison_table.to_excel(writer, index=False, sheet_name='Comparison')
                period1_data.to_excel(writer, index=False, sheet_name='Period 1 Data')
                period2_data.to_excel(writer, index=False, sheet_name='Period 2 Data')
            comparison_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Comparison Data",
                data=comparison_buffer,
                file_name=f"Period_Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Download charts
            col1, col2 = st.columns(2)
            with col1:
                try:
                    # Create copy with filter annotation
                    fig_comp_with_filters = px.bar(
                        comparison_data,
                        x='Period',
                        y='Sales Qty',
                        title='Period Comparison',
                        text='Sales Qty'
                    )
                    fig_comp_with_filters.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                    fig_comp_with_filters.update_layout(
                        height=400,
                        margin=dict(b=100),
                        title_text=f"Period Comparison<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                        title_x=0.5
                    )
                    fig_comp_with_filters = add_filter_annotation(fig_comp_with_filters, filter_text)
                    comparison_img = fig_comp_with_filters.to_image(format="png")
                    st.download_button(
                        label="📥 Download Comparison Chart",
                        data=comparison_img,
                        file_name=f"Period_Comparison_Chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except Exception as e:
                    st.info("Chart export available via interactive chart above")
            with col2:
                try:
                    # Create copy with filter annotation
                    fig_monthly_with_filters = px.line(
                        monthly_comparison,
                        x='Mth-yr',
                        y='Sales Qty',
                        color='Period',
                        title='Monthly Breakdown by Period',
                        markers=True
                    )
                    fig_monthly_with_filters.update_layout(
                        height=400,
                        margin=dict(b=100),
                        title_text=f"Monthly Breakdown by Period<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                        title_x=0.5
                    )
                    fig_monthly_with_filters = add_filter_annotation(fig_monthly_with_filters, filter_text)
                    monthly_img = fig_monthly_with_filters.to_image(format="png")
                    st.download_button(
                        label="📥 Download Monthly Breakdown Chart",
                        data=monthly_img,
                        file_name=f"Monthly_Breakdown_Chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                        mime="image/png"
                    )
                except Exception as e:
                    st.info("Chart export available via interactive chart above")
        else:
            st.warning("Please select valid date ranges for both periods.")
    
    with tab3:
        st.subheader("Region Comparison")
        
        # Single time period selection for Region Comparison
        st.markdown("### Time Period Selection")
        col1, col2 = st.columns(2)
        with col1:
            region_start = st.date_input(
                "Start Date",
                value=actual_min,
                min_value=picker_min,
                max_value=picker_max,
                key="region_start"
            )
        with col2:
            region_end = st.date_input(
                "End Date",
                value=actual_max,
                min_value=picker_min,
                max_value=picker_max,
                key="region_end"
            )
        
        # Validate dates
        if region_start < min_date:
            region_start = min_date
        if region_end > max_date:
            region_end = max_date
        
        # Filter data by time period
        region_start_dt = pd.to_datetime(region_start)
        region_end_dt = pd.to_datetime(region_end)
        region_data = filtered_df[
            (filtered_df['Mth-yr'] >= region_start_dt) &
            (filtered_df['Mth-yr'] <= region_end_dt)
        ]
        
        if len(region_data) == 0:
            st.warning("No data available for the selected time period.")
        else:
            # Region 1 and Region 2 selection
            st.markdown("### Region Selection")
            available_regions = sorted(region_data['Region'].unique())
            col1, col2 = st.columns(2)
            
            with col1:
                region1 = st.selectbox(
                    "Region 1",
                    options=available_regions,
                    key="region1"
                )
            with col2:
                region2 = st.selectbox(
                    "Region 2",
                    options=available_regions,
                    key="region2"
                )
            
            if region1 == region2:
                st.warning("Please select two different regions for comparison.")
            else:
                # Process region comparison with progress
                tab3_progress = st.progress(0)
                tab3_status = st.empty()
                tab3_status.text("📊 Processing region comparison...")
                tab3_progress.progress(30)
                
                # Filter data for each region
                region1_data = region_data[region_data['Region'] == region1]
                region2_data = region_data[region_data['Region'] == region2]
                
                # Calculate totals
                region1_total = region1_data['Sales Qty'].sum()
                region2_total = region2_data['Sales Qty'].sum()
                change = region2_total - region1_total
                change_pct = (change / region1_total * 100) if region1_total > 0 else 0
                
                tab3_progress.progress(60)
                
                # Display metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(f"{region1} Total", f"{region1_total:,.0f}")
                with col2:
                    st.metric(f"{region2} Total", f"{region2_total:,.0f}")
                with col3:
                    st.metric("Change", f"{change:,.0f}", f"{change_pct:.1f}%")
                
                # Comparison chart with Customer Type breakdown
                # Get customer type breakdown for each region
                region1_by_customer = region1_data.groupby('Customer Type')['Sales Qty'].sum().reset_index()
                region1_by_customer['Region'] = region1
                region2_by_customer = region2_data.groupby('Customer Type')['Sales Qty'].sum().reset_index()
                region2_by_customer['Region'] = region2
                
                # Combine and add total row
                comparison_by_customer = pd.concat([region1_by_customer, region2_by_customer])
                
                # Add total rows
                total_rows = pd.DataFrame({
                    'Customer Type': ['Total'],
                    'Sales Qty': [region1_total],
                    'Region': [region1]
                })
                comparison_by_customer = pd.concat([comparison_by_customer, total_rows])
                
                total_rows2 = pd.DataFrame({
                    'Customer Type': ['Total'],
                    'Sales Qty': [region2_total],
                    'Region': [region2]
                })
                comparison_by_customer = pd.concat([comparison_by_customer, total_rows2])
                
                tab3_progress.progress(80)
                tab3_status.text("📊 Generating charts...")
                
                # Format time period for title
                time_period_text = f"Period: {region_start.strftime('%b %d, %Y')} to {region_end.strftime('%b %d, %Y')}"
                
                fig_region_comp = px.bar(
                    comparison_by_customer,
                    x='Region',
                    y='Sales Qty',
                    color='Customer Type',
                    title=f'Region Comparison: {region1} vs {region2}',
                    text='Sales Qty',
                    barmode='group'
                )
                fig_region_comp.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                # Add time period to title
                fig_region_comp.update_layout(
                    height=400,
                    title_text=f"Region Comparison: {region1} vs {region2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                    title_x=0.5
                )
                st.plotly_chart(fig_region_comp, use_container_width=True)
                
                # Monthly breakdown
                region1_monthly = region1_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
                region1_monthly['Region'] = region1
                region2_monthly = region2_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
                region2_monthly['Region'] = region2
                
                monthly_comparison = pd.concat([region1_monthly, region2_monthly])
                
                fig_monthly_region = px.line(
                    monthly_comparison,
                    x='Mth-yr',
                    y='Sales Qty',
                    color='Region',
                    title=f'Monthly Breakdown: {region1} vs {region2}',
                    markers=True
                )
                # Add time period to title
                fig_monthly_region.update_layout(
                    height=400,
                    title_text=f"Monthly Breakdown: {region1} vs {region2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                    title_x=0.5
                )
                st.plotly_chart(fig_monthly_region, use_container_width=True)
                
                tab3_progress.progress(100)
                tab3_status.text("✅ Data processed!")
                import time
                time.sleep(0.1)
                tab3_progress.empty()
                tab3_status.empty()
                
                # Download charts
                col1, col2 = st.columns(2)
                with col1:
                    try:
                        fig_region_comp_with_filters = px.bar(
                            comparison_by_customer,
                            x='Region',
                            y='Sales Qty',
                            color='Customer Type',
                            title=f'Region Comparison: {region1} vs {region2}',
                            text='Sales Qty',
                            barmode='group'
                        )
                        fig_region_comp_with_filters.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                        fig_region_comp_with_filters.update_layout(
                            height=400,
                            margin=dict(b=100),
                            title_text=f"Region Comparison: {region1} vs {region2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                            title_x=0.5
                        )
                        fig_region_comp_with_filters = add_filter_annotation(fig_region_comp_with_filters, filter_text)
                        region_comp_img = fig_region_comp_with_filters.to_image(format="png")
                        st.download_button(
                            label="📥 Download Region Comparison Chart",
                            data=region_comp_img,
                            file_name=f"Region_Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                            mime="image/png"
                        )
                    except Exception as e:
                        st.info("Chart export available via interactive chart above")
                with col2:
                    try:
                        fig_monthly_region_with_filters = px.line(
                            monthly_comparison,
                            x='Mth-yr',
                            y='Sales Qty',
                            color='Region',
                            title=f'Monthly Breakdown: {region1} vs {region2}',
                            markers=True
                        )
                        fig_monthly_region_with_filters.update_layout(
                            height=400,
                            margin=dict(b=100),
                            title_text=f"Monthly Breakdown: {region1} vs {region2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                            title_x=0.5
                        )
                        fig_monthly_region_with_filters = add_filter_annotation(fig_monthly_region_with_filters, filter_text)
                        monthly_region_img = fig_monthly_region_with_filters.to_image(format="png")
                        st.download_button(
                            label="📥 Download Monthly Breakdown Chart",
                            data=monthly_region_img,
                            file_name=f"Region_Monthly_Breakdown_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                            mime="image/png"
                        )
                    except Exception as e:
                        st.info("Chart export available via interactive chart above")
    
    with tab4:
        st.subheader("Township Comparison")
        
        # Single time period selection for Township Comparison
        st.markdown("### Time Period Selection")
        col1, col2 = st.columns(2)
        with col1:
            township_start = st.date_input(
                "Start Date",
                value=actual_min,
                min_value=picker_min,
                max_value=picker_max,
                key="township_start"
            )
        with col2:
            township_end = st.date_input(
                "End Date",
                value=actual_max,
                min_value=picker_min,
                max_value=picker_max,
                key="township_end"
            )
        
        # Validate dates
        if township_start < min_date:
            township_start = min_date
        if township_end > max_date:
            township_end = max_date
        
        # Filter data by time period
        township_start_dt = pd.to_datetime(township_start)
        township_end_dt = pd.to_datetime(township_end)
        township_data = filtered_df[
            (filtered_df['Mth-yr'] >= township_start_dt) &
            (filtered_df['Mth-yr'] <= township_end_dt)
        ]
        
        if len(township_data) == 0:
            st.warning("No data available for the selected time period.")
        else:
            # Township 1 and Township 2 selection
            st.markdown("### Township Selection")
            available_townships = sorted(township_data['Township'].unique())
            col1, col2 = st.columns(2)
            
            with col1:
                township1 = st.selectbox(
                    "Township 1",
                    options=available_townships,
                    key="township1"
                )
            with col2:
                township2 = st.selectbox(
                    "Township 2",
                    options=available_townships,
                    key="township2"
                )
            
            if township1 == township2:
                st.warning("Please select two different townships for comparison.")
            else:
                # Process township comparison with progress
                tab4_progress = st.progress(0)
                tab4_status = st.empty()
                tab4_status.text("📊 Processing township comparison...")
                tab4_progress.progress(30)
                
                # Filter data for each township
                township1_data = township_data[township_data['Township'] == township1]
                township2_data = township_data[township_data['Township'] == township2]
                
                # Calculate totals
                township1_total = township1_data['Sales Qty'].sum()
                township2_total = township2_data['Sales Qty'].sum()
                change = township2_total - township1_total
                change_pct = (change / township1_total * 100) if township1_total > 0 else 0
                
                tab4_progress.progress(60)
                
                # Display metrics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(f"{township1} Total", f"{township1_total:,.0f}")
                with col2:
                    st.metric(f"{township2} Total", f"{township2_total:,.0f}")
                with col3:
                    st.metric("Change", f"{change:,.0f}", f"{change_pct:.1f}%")
                
                # Comparison chart with Customer Type breakdown
                # Get customer type breakdown for each township
                township1_by_customer = township1_data.groupby('Customer Type')['Sales Qty'].sum().reset_index()
                township1_by_customer['Township'] = township1
                township2_by_customer = township2_data.groupby('Customer Type')['Sales Qty'].sum().reset_index()
                township2_by_customer['Township'] = township2
                
                # Combine and add total row
                comparison_by_customer = pd.concat([township1_by_customer, township2_by_customer])
                
                # Add total rows
                total_rows = pd.DataFrame({
                    'Customer Type': ['Total'],
                    'Sales Qty': [township1_total],
                    'Township': [township1]
                })
                comparison_by_customer = pd.concat([comparison_by_customer, total_rows])
                
                total_rows2 = pd.DataFrame({
                    'Customer Type': ['Total'],
                    'Sales Qty': [township2_total],
                    'Township': [township2]
                })
                comparison_by_customer = pd.concat([comparison_by_customer, total_rows2])
                
                tab4_progress.progress(80)
                tab4_status.text("📊 Generating charts...")
                
                # Format time period for title
                time_period_text = f"Period: {township_start.strftime('%b %d, %Y')} to {township_end.strftime('%b %d, %Y')}"
                
                fig_township_comp = px.bar(
                    comparison_by_customer,
                    x='Township',
                    y='Sales Qty',
                    color='Customer Type',
                    title=f'Township Comparison: {township1} vs {township2}',
                    text='Sales Qty',
                    barmode='group'
                )
                fig_township_comp.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                # Add time period to title
                fig_township_comp.update_layout(
                    height=400,
                    title_text=f"Township Comparison: {township1} vs {township2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                    title_x=0.5
                )
                st.plotly_chart(fig_township_comp, use_container_width=True)
                
                # Monthly breakdown
                township1_monthly = township1_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
                township1_monthly['Township'] = township1
                township2_monthly = township2_data.groupby('Mth-yr')['Sales Qty'].sum().reset_index()
                township2_monthly['Township'] = township2
                
                monthly_comparison = pd.concat([township1_monthly, township2_monthly])
                
                fig_monthly_township = px.line(
                    monthly_comparison,
                    x='Mth-yr',
                    y='Sales Qty',
                    color='Township',
                    title=f'Monthly Breakdown: {township1} vs {township2}',
                    markers=True
                )
                # Add time period to title
                fig_monthly_township.update_layout(
                    height=400,
                    title_text=f"Monthly Breakdown: {township1} vs {township2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                    title_x=0.5
                )
                st.plotly_chart(fig_monthly_township, use_container_width=True)
                
                tab4_progress.progress(100)
                tab4_status.text("✅ Data processed!")
                import time
                time.sleep(0.1)
                tab4_progress.empty()
                tab4_status.empty()
                
                # Download charts
                col1, col2 = st.columns(2)
                with col1:
                    try:
                        fig_township_comp_with_filters = px.bar(
                            comparison_by_customer,
                            x='Township',
                            y='Sales Qty',
                            color='Customer Type',
                            title=f'Township Comparison: {township1} vs {township2}',
                            text='Sales Qty',
                            barmode='group'
                        )
                        fig_township_comp_with_filters.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                        fig_township_comp_with_filters.update_layout(
                            height=400,
                            margin=dict(b=100),
                            title_text=f"Township Comparison: {township1} vs {township2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                            title_x=0.5
                        )
                        fig_township_comp_with_filters = add_filter_annotation(fig_township_comp_with_filters, filter_text)
                        township_comp_img = fig_township_comp_with_filters.to_image(format="png")
                        st.download_button(
                            label="📥 Download Township Comparison Chart",
                            data=township_comp_img,
                            file_name=f"Township_Comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                            mime="image/png"
                        )
                    except Exception as e:
                        st.info("Chart export available via interactive chart above")
                with col2:
                    try:
                        fig_monthly_township_with_filters = px.line(
                            monthly_comparison,
                            x='Mth-yr',
                            y='Sales Qty',
                            color='Township',
                            title=f'Monthly Breakdown: {township1} vs {township2}',
                            markers=True
                        )
                        fig_monthly_township_with_filters.update_layout(
                            height=400,
                            margin=dict(b=100),
                            title_text=f"Monthly Breakdown: {township1} vs {township2}<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                            title_x=0.5
                        )
                        fig_monthly_township_with_filters = add_filter_annotation(fig_monthly_township_with_filters, filter_text)
                        monthly_township_img = fig_monthly_township_with_filters.to_image(format="png")
                        st.download_button(
                            label="📥 Download Monthly Breakdown Chart",
                            data=monthly_township_img,
                            file_name=f"Township_Monthly_Breakdown_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                            mime="image/png"
                        )
                    except Exception as e:
                        st.info("Chart export available via interactive chart above")
    
    with tab4:
        st.subheader("Product Analysis")
        
        # Single time period selection for Product Analysis
        st.markdown("### Time Period Selection")
        col1, col2 = st.columns(2)
        with col1:
            product_start = st.date_input(
                "Start Date",
                value=actual_min,
                min_value=picker_min,
                max_value=picker_max,
                key="product_start"
            )
        with col2:
            product_end = st.date_input(
                "End Date",
                value=actual_max,
                min_value=picker_min,
                max_value=picker_max,
                key="product_end"
            )
        
        # Validate dates
        if product_start < min_date:
            product_start = min_date
        if product_end > max_date:
            product_end = max_date
        
        # Filter data by time period
        product_start_dt = pd.to_datetime(product_start)
        product_end_dt = pd.to_datetime(product_end)
        product_data = filtered_df[
            (filtered_df['Mth-yr'] >= product_start_dt) &
            (filtered_df['Mth-yr'] <= product_end_dt)
        ]
        
        if len(product_data) == 0:
            st.warning("No data available for the selected time period.")
        else:
            # Process product analysis with progress
            tab4_progress = st.progress(0)
            tab4_status = st.empty()
            tab4_status.text("📊 Processing product data...")
            tab4_progress.progress(30)
            
            # Product sales
            product_sales = product_data.groupby('Product')['Sales Qty'].sum().reset_index().sort_values('Sales Qty', ascending=False)
            tab4_progress.progress(60)
            
            # Format time period for title
            time_period_text = f"Period: {product_start.strftime('%b %d, %Y')} to {product_end.strftime('%b %d, %Y')}"
            
            fig_product = px.bar(
                product_sales,
                x='Product',
                y='Sales Qty',
                title='Sales by Product',
                text='Sales Qty'
            )
            fig_product.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
            # Add time period to title
            fig_product.update_layout(
                height=400,
                xaxis_tickangle=-45,
                title_text=f"Sales by Product<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                title_x=0.5
            )
            st.plotly_chart(fig_product, use_container_width=True)
            
            # Product trend over time
            product_trend = product_data.groupby(['Mth-yr', 'Product'])['Sales Qty'].sum().reset_index()
            tab4_progress.progress(70)
            
            fig_product_trend = px.line(
                product_trend,
                x='Mth-yr',
                y='Sales Qty',
                color='Product',
                title='Product Sales Trend Over Time',
                markers=True
            )
            # Add time period to title
            fig_product_trend.update_layout(
                height=400,
                title_text=f"Product Sales Trend Over Time<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                title_x=0.5
            )
            st.plotly_chart(fig_product_trend, use_container_width=True)
            
            # Product by Customer Type
            product_customer = product_data.groupby(['Product', 'Customer Type'])['Sales Qty'].sum().reset_index()
            
            tab4_progress.progress(90)
            tab4_status.text("📊 Generating charts...")
            
            fig_product_customer = px.bar(
                product_customer,
                x='Product',
                y='Sales Qty',
                color='Customer Type',
                title='Sales by Product and Customer Type',
                barmode='group'
            )
            # Add time period to title
            fig_product_customer.update_layout(
                height=400,
                xaxis_tickangle=-45,
                title_text=f"Sales by Product and Customer Type<br><sub style='font-size: 0.7em;'>{time_period_text}</sub>",
                title_x=0.5
            )
            st.plotly_chart(fig_product_customer, use_container_width=True)
            
            tab4_progress.progress(100)
            tab4_status.text("✅ Data processed!")
            import time
            time.sleep(0.1)
            tab4_progress.empty()
            tab4_status.empty()
            
            # Download product data
            product_buffer = io.BytesIO()
            with pd.ExcelWriter(product_buffer, engine='xlsxwriter') as writer:
                product_sales.to_excel(writer, index=False, sheet_name='Product Sales')
                product_trend.to_excel(writer, index=False, sheet_name='Product Trend')
                product_customer.to_excel(writer, index=False, sheet_name='Product-Customer')
            product_buffer.seek(0)
            
            st.download_button(
                label="📥 Download Product Analysis Data",
                data=product_buffer,
                file_name=f"Product_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()


