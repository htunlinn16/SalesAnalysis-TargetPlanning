# Sales Target & Analysis App

A comprehensive Streamlit application for sales target planning and sales trend analysis.

## Features

### 🎯 Target Planning
- **Template Download**: Download a pre-formatted Excel template for sales data entry
- **Data Upload**: Upload sales data (Excel or CSV format)
- **AMS Calculation**: Automatically calculates Average Monthly Sales (AMS) for the last 6 months
  - Filters out months with sales < 20% of AMS (to account for product shortages)
- **Target Calculation**: Calculate sales targets based on percentage increase on AMS
- **Multi-select Filters**: Filter by Product, Customer Type, Township, and Region
- **Data Downloads**: Download calculated AMS and Target data as Excel files

### 📈 Sales Analysis
- **Trend Analysis**: Visualize sales trends over time with interactive charts
- **Period Comparison**: Compare sales between two different time periods
- **Region/Township Comparison**: Analyze sales by region and township with heatmaps
- **Product Analysis**: Deep dive into product performance
- **Interactive Charts**: Multiple chart types including line, bar, and heatmap charts
- **Download Capabilities**: Download all charts and comparison tables

## Installation

1. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit app:
```bash
streamlit run app.py
```

2. The app will open in your browser at `http://localhost:8501`

## Data Template

The template includes the following columns:
- **Mth-yr**: Month and year - accepts multiple date formats including:
  - `Jan-2024`, `Feb-2024` (mmm-yr format)
  - `2024-01-01`, `2024/01/01` (ISO date format)
  - `01/2024`, `01-2024` (month/year format)
  - `January 2024`, `Jan 2024` (month name format)
  - `2024-01`, `2024/01` (year-month format)
  - Excel date serial numbers
  - Any other format that contains month and year information
- **Product**: Product name
- **Customer Type**: Type of customer (e.g., Retail, Wholesale, Corporate)
- **Township**: Township name
- **Region**: Region name
- **Sales Qty**: Sales quantity (numeric)

The template comes pre-filled with sample data for the last 12 months to help you understand the format. The app automatically parses and standardizes all date formats internally.

## Features

### Design
- **Modern, Clean Interface**: Beautiful and intuitive design
- **Mobile Responsive**: Fully optimized for iPad and mobile devices
- **Custom Styling**: Professional appearance with smooth animations and transitions

## How It Works

### AMS Calculation
1. Identifies the last 6 months from the maximum date in the data
2. Groups data by Product, Customer Type, Township, and Region
3. Calculates initial AMS from all months
4. Filters out months where sales < 20% of initial AMS
5. Recalculates AMS using only the filtered months

### Target Calculation
Targets are calculated as: `Target Qty = AMS × (1 + Percentage Increase / 100)`

## Notes

- The app uses session state to maintain data across page navigation
- All charts are interactive and can be zoomed, panned, and exported
- Excel files are generated using xlsxwriter for better formatting


