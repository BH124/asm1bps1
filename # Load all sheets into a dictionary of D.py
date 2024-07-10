import pandas as pd

# Load the Excel file
file_path = '/mnt/data/Book1.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load all sheets into a dictionary of DataFrames
sheets_data = {sheet: excel_data.parse(sheet) for sheet in excel_data.sheet_names}

# Extract DataFrames
customer_df = sheets_data['Customer']
sale_df = sheets_data['Sale']
product_detail_df = sheets_data['ProductDetail']
product_group_df = sheets_data['ProductGroup']
market_trend_df = sheets_data['MarketTrend']

# Assuming we have columns to relate these tables

# If ProductGroupID exists in ProductDetail, add it
if 'ProductGroupID' not in product_detail_df.columns:
    product_detail_df = product_detail_df.merge(
        product_group_df[['ProductGroupID']],
        left_on='ProductID',
        right_index=True,
        how='left'
    )

# Save each DataFrame as a CSV file
def save_csv(df, name):
    path = f'/mnt/data/{name}.csv'
    df.to_csv(path, index=False)
    return path

csv_paths = {
    'Customer': save_csv(customer_df, 'Customer'),
    'Sale': save_csv(sale_df, 'Sale'),
    'ProductDetail': save_csv(product_detail_df, 'ProductDetail'),
    'ProductGroup': save_csv(product_group_df, 'ProductGroup'),
    'MarketTrend': save_csv(market_trend_df, 'MarketTrend')
}

csv_paths
