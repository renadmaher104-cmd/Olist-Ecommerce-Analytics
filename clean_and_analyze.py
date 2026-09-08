"""
Olist E-Commerce Data Cleaning, Analytics & Executive Excel Workbook Generator
=============================================================================
This script performs:
1. Data Ingestion & Cleaning of 9 relational datasets.
2. Feature engineering (Logistics lead times, delay flags, date dimensions).
3. Product category English translation & imputation.
4. Aggregation of 5 business reports (Sales Trend, Category Insights, Regional Logistics, Payment Methods, CSAT Reviews).
5. Generation of an executive-grade, multi-tab Excel Workbook with KPI cards and native visual charts using XlsxWriter.
6. Export of clean datasets to CSV.
"""

import os
import sys
import numpy as np
import pandas as pd
import xlsxwriter

def load_and_clean_data(data_dir):
    print(">>> Step 1: Loading raw datasets...")
    orders = pd.read_csv(os.path.join(data_dir, 'olist_orders_dataset.csv'))
    items = pd.read_csv(os.path.join(data_dir, 'olist_order_items_dataset.csv'))
    payments = pd.read_csv(os.path.join(data_dir, 'olist_order_payments_dataset.csv'))
    reviews = pd.read_csv(os.path.join(data_dir, 'olist_order_reviews_dataset.csv'))
    products = pd.read_csv(os.path.join(data_dir, 'olist_products_dataset.csv'))
    customers = pd.read_csv(os.path.join(data_dir, 'olist_customers_dataset.csv'))
    sellers = pd.read_csv(os.path.join(data_dir, 'olist_sellers_dataset.csv'))
    translations = pd.read_csv(os.path.join(data_dir, 'product_category_name_translation.csv'))

    print(">>> Step 2: Cleaning and standardizing dates...")
    date_cols = [
        'order_purchase_timestamp',
        'order_approved_at',
        'order_delivered_carrier_date',
        'order_delivered_customer_date',
        'order_estimated_delivery_date'
    ]
    for col in date_cols:
        orders[col] = pd.to_datetime(orders[col], errors='coerce')

    reviews['review_creation_date'] = pd.to_datetime(reviews['review_creation_date'], errors='coerce')
    reviews['review_answer_timestamp'] = pd.to_datetime(reviews['review_answer_timestamp'], errors='coerce')

    # Date feature engineering
    orders['purchase_year'] = orders['order_purchase_timestamp'].dt.year
    orders['purchase_month'] = orders['order_purchase_timestamp'].dt.month
    orders['purchase_year_month'] = orders['order_purchase_timestamp'].dt.to_period('M').astype(str)
    orders['purchase_quarter'] = orders['order_purchase_timestamp'].dt.to_period('Q').astype(str)
    orders['purchase_day_name'] = orders['order_purchase_timestamp'].dt.day_name()
    orders['purchase_hour'] = orders['order_purchase_timestamp'].dt.hour

    # Logistics calculations (in days)
    orders['actual_delivery_days'] = (
        orders['order_delivered_customer_date'] - orders['order_purchase_timestamp']
    ).dt.total_seconds() / (24 * 3600)

    orders['estimated_delivery_days'] = (
        orders['order_estimated_delivery_date'] - orders['order_purchase_timestamp']
    ).dt.total_seconds() / (24 * 3600)

    orders['delivery_delay_days'] = (
        orders['order_delivered_customer_date'] - orders['order_estimated_delivery_date']
    ).dt.total_seconds() / (24 * 3600)

    orders['is_delayed'] = np.where(orders['delivery_delay_days'] > 0, 1, 0)
    orders['delivery_status'] = np.where(orders['is_delayed'] == 1, 'Late', 'On-Time / Early')

    print(">>> Step 3: Cleaning products and category translations...")
    products = products.merge(translations, on='product_category_name', how='left')
    products['product_category_name_english'] = products['product_category_name_english'].fillna('other')
    products['product_category_clean'] = products['product_category_name_english'].str.replace('_', ' ').str.title()
    
    # Impute missing dimensions with median
    for dim_col in ['product_weight_g', 'product_length_cm', 'product_height_cm', 'product_width_cm', 'product_photos_qty']:
        median_val = products[dim_col].median()
        products[dim_col] = products[dim_col].fillna(median_val)

    print(">>> Step 4: Cleaning payments dataset...")
    # Filter or replace 'not_defined' payment types
    payments = payments[payments['payment_type'] != 'not_defined'].copy()
    payments['payment_type_clean'] = payments['payment_type'].str.replace('_', ' ').str.title()

    # Aggregate payments to order level
    order_payments_agg = payments.groupby('order_id').agg(
        total_payment_value=('payment_value', 'sum'),
        payment_installments_max=('payment_installments', 'max'),
        payment_methods_count=('payment_type', 'nunique'),
        primary_payment_type=('payment_type_clean', lambda x: x.mode()[0] if not x.empty else 'Unknown')
    ).reset_index()

    print(">>> Step 5: Cleaning and deduplicating reviews...")
    # Orders can have multiple reviews; aggregate by order
    order_reviews_agg = reviews.groupby('order_id').agg(
        avg_review_score=('review_score', 'mean'),
        min_review_score=('review_score', 'min'),
        max_review_score=('review_score', 'max'),
        review_count=('review_score', 'count'),
        has_comment_title=('review_comment_title', lambda x: x.notnull().any().astype(int)),
        has_comment_message=('review_comment_message', lambda x: x.notnull().any().astype(int))
    ).reset_index()

    print(">>> Step 6: Merging items with product, customer, seller details...")
    # Items aggregated to order level and item-level master
    items_enriched = items.merge(
        products[['product_id', 'product_category_clean', 'product_weight_g']], 
        on='product_id', 
        how='left'
    )
    items_enriched = items_enriched.merge(
        sellers[['seller_id', 'seller_city', 'seller_state']], 
        on='seller_id', 
        how='left'
    )

    order_items_agg = items.groupby('order_id').agg(
        total_items_price=('price', 'sum'),
        total_freight_value=('freight_value', 'sum'),
        items_count=('order_item_id', 'count'),
        sellers_count=('seller_id', 'nunique'),
        products_count=('product_id', 'nunique')
    ).reset_index()
    order_items_agg['total_order_value'] = order_items_agg['total_items_price'] + order_items_agg['total_freight_value']

    print(">>> Step 7: Creating Master Orders Dataset...")
    master_orders = orders.merge(customers, on='customer_id', how='left')
    master_orders = master_orders.merge(order_items_agg, on='order_id', how='left')
    master_orders = master_orders.merge(order_payments_agg, on='order_id', how='left')
    master_orders = master_orders.merge(order_reviews_agg, on='order_id', how='left')

    # Regional mapping for Brazilian states
    region_map = {
        'SP': 'Southeast', 'RJ': 'Southeast', 'MG': 'Southeast', 'ES': 'Southeast',
        'RS': 'South', 'PR': 'South', 'SC': 'South',
        'BA': 'Northeast', 'PE': 'Northeast', 'CE': 'Northeast', 'MA': 'Northeast', 
        'PB': 'Northeast', 'RN': 'Northeast', 'AL': 'Northeast', 'SE': 'Northeast', 'PI': 'Northeast',
        'GO': 'Central-West', 'MT': 'Central-West', 'MS': 'Central-West', 'DF': 'Central-West',
        'PA': 'North', 'AM': 'North', 'RO': 'North', 'TO': 'North', 'AC': 'North', 'AP': 'North', 'RR': 'North'
    }
    master_orders['customer_region'] = master_orders['customer_state'].map(region_map).fillna('Other')

    # Export Clean CSVs
    output_clean_dir = os.path.join(os.path.dirname(data_dir), 'cleaned_data')
    os.makedirs(output_clean_dir, exist_ok=True)
    
    print(f">>> Exporting clean CSVs to {output_clean_dir}...")
    master_orders.to_csv(os.path.join(output_clean_dir, 'clean_orders_master.csv'), index=False)
    products.to_csv(os.path.join(output_clean_dir, 'clean_products.csv'), index=False)
    items_enriched.to_csv(os.path.join(output_clean_dir, 'clean_order_items.csv'), index=False)
    payments.to_csv(os.path.join(output_clean_dir, 'clean_payments.csv'), index=False)

    return {
        'orders': orders,
        'master_orders': master_orders,
        'items_enriched': items_enriched,
        'payments': payments,
        'reviews': reviews,
        'products': products,
        'customers': customers
    }

def generate_reports_data(data_dict):
    print(">>> Step 8: Generating 5 Analytical Reports...")
    master = data_dict['master_orders']
    items_enriched = data_dict['items_enriched']
    payments = data_dict['payments']
    reviews = data_dict['reviews']

    # Filter to delivered or valid orders for primary sales metrics
    delivered = master[master['order_status'] == 'delivered'].copy()

    # --- REPORT 1: Sales & Revenue Trend (Monthly) ---
    # Filter 2017-01 to 2018-08 for clean full operational months
    monthly_sales = delivered[
        (delivered['purchase_year_month'] >= '2017-01') & 
        (delivered['purchase_year_month'] <= '2018-08')
    ].groupby('purchase_year_month').agg(
        total_orders=('order_id', 'count'),
        total_revenue=('total_order_value', 'sum'),
        items_revenue=('total_items_price', 'sum'),
        freight_revenue=('total_freight_value', 'sum'),
        avg_order_value=('total_order_value', 'mean'),
        avg_delivery_days=('actual_delivery_days', 'mean'),
        on_time_rate=('is_delayed', lambda x: 1.0 - x.mean())
    ).reset_index()

    monthly_sales['mom_revenue_growth'] = monthly_sales['total_revenue'].pct_change().fillna(0)
    monthly_sales['freight_ratio'] = monthly_sales['freight_revenue'] / monthly_sales['total_revenue']

    # --- REPORT 2: Product Category Insights ---
    # Merge item with order status
    items_with_status = items_enriched.merge(
        master[['order_id', 'order_status', 'avg_review_score']], 
        on='order_id', 
        how='left'
    )
    category_summary = items_with_status[items_with_status['order_status'] == 'delivered'].groupby('product_category_clean').agg(
        total_items_sold=('order_item_id', 'count'),
        total_revenue=('price', 'sum'),
        total_freight=('freight_value', 'sum'),
        avg_item_price=('price', 'mean'),
        avg_review_score=('avg_review_score', 'mean')
    ).reset_index()

    category_summary['share_of_revenue'] = category_summary['total_revenue'] / category_summary['total_revenue'].sum()
    category_summary = category_summary.sort_values(by='total_revenue', ascending=False).reset_index(drop=True)

    # --- REPORT 3: Regional Logistics & Delivery Performance ---
    regional_summary = delivered.groupby(['customer_region', 'customer_state']).agg(
        total_orders=('order_id', 'count'),
        total_revenue=('total_order_value', 'sum'),
        avg_delivery_days=('actual_delivery_days', 'mean'),
        avg_estimated_days=('estimated_delivery_days', 'mean'),
        avg_delay_days=('delivery_delay_days', lambda x: x[x > 0].mean() if (x > 0).any() else 0),
        on_time_rate=('is_delayed', lambda x: 1.0 - x.mean()),
        avg_freight_value=('total_freight_value', 'mean')
    ).reset_index()

    regional_summary['share_of_orders'] = regional_summary['total_orders'] / regional_summary['total_orders'].sum()
    regional_summary = regional_summary.sort_values(by='total_orders', ascending=False).reset_index(drop=True)

    # --- REPORT 4: Payment Methods & Financing ---
    payment_summary = payments.groupby('payment_type_clean').agg(
        transaction_count=('payment_value', 'count'),
        total_payment_value=('payment_value', 'sum'),
        avg_transaction_value=('payment_value', 'mean'),
        avg_installments=('payment_installments', 'mean'),
        max_installments=('payment_installments', 'max')
    ).reset_index()

    payment_summary['share_of_gmv'] = payment_summary['total_payment_value'] / payment_summary['total_payment_value'].sum()
    payment_summary['share_of_transactions'] = payment_summary['transaction_count'] / payment_summary['transaction_count'].sum()
    payment_summary = payment_summary.sort_values(by='total_payment_value', ascending=False).reset_index(drop=True)

    # --- REPORT 5: Customer Satisfaction & Review Analysis ---
    review_summary = reviews.groupby('review_score').agg(
        review_count=('review_id', 'count'),
        with_comment_message=('review_comment_message', lambda x: x.notnull().sum()),
        with_comment_title=('review_comment_title', lambda x: x.notnull().sum())
    ).reset_index()

    review_summary['percentage_of_reviews'] = review_summary['review_count'] / review_summary['review_count'].sum()
    review_summary['comment_rate'] = review_summary['with_comment_message'] / review_summary['review_count']
    review_summary = review_summary.sort_values(by='review_score', ascending=False).reset_index(drop=True)

    # Cross-tab of Delivery Status vs Avg Review Score
    delivery_vs_review = delivered.groupby('delivery_status').agg(
        order_count=('order_id', 'count'),
        avg_review_score=('avg_review_score', 'mean'),
        avg_actual_delivery_days=('actual_delivery_days', 'mean')
    ).reset_index()

    return {
        'monthly_sales': monthly_sales,
        'category_summary': category_summary,
        'regional_summary': regional_summary,
        'payment_summary': payment_summary,
        'review_summary': review_summary,
        'delivery_vs_review': delivery_vs_review
    }

def create_excel_workbook(data_dict, reports_dict, output_filepath):
    print(f">>> Step 9: Creating Master Excel Workbook with XlsxWriter at {output_filepath}...")
    workbook = xlsxwriter.Workbook(output_filepath, {'nan_inf_to_errors': True})

    # ------------------ FORMAT DEFINITIONS ------------------
    NAVY = '#0F172A'
    BLUE_HEADER = '#1E3A8A'
    SLATE_LIGHT = '#F1F5F9'
    BORDER_COLOR = '#CBD5E1'

    fmt_title = workbook.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Segoe UI', 'font_color': '#FFFFFF',
        'bg_color': NAVY, 'align': 'left', 'valign': 'vcenter', 'indent': 1
    })
    fmt_subtitle = workbook.add_format({
        'font_size': 10, 'font_name': 'Segoe UI', 'font_color': '#94A3B8',
        'bg_color': NAVY, 'align': 'left', 'valign': 'vcenter', 'indent': 1
    })
    fmt_kpi_card_title = workbook.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#475569',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter',
        'top': 1, 'left': 1, 'right': 1, 'top_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    })
    fmt_kpi_card_val_curr = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_name': 'Segoe UI', 'font_color': '#0284C7',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter', 'num_format': 'R$ #,##0',
        'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    })
    fmt_kpi_card_val_num = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_name': 'Segoe UI', 'font_color': '#0F172A',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0',
        'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    })
    fmt_kpi_card_val_pct = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_name': 'Segoe UI', 'font_color': '#10B981',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0%',
        'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    })
    fmt_sec_header = workbook.add_format({
        'bold': True, 'font_size': 11, 'font_name': 'Segoe UI', 'font_color': '#1E293B',
        'bottom': 2, 'bottom_color': '#3B82F6', 'valign': 'vcenter'
    })
    fmt_tbl_header = workbook.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Segoe UI', 'font_color': '#FFFFFF',
        'bg_color': BLUE_HEADER, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'border_color': '#1E293B'
    })
    fmt_cell_text = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_text_center = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_num = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': '#,##0', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_curr = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': 'R$ #,##0.00', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_pct = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': '0.0%', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_dec = workbook.add_format({
        'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': '0.00', 'align': 'right', 'valign': 'vcenter',
        'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_total = workbook.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Segoe UI', 'bg_color': SLATE_LIGHT, 'valign': 'vcenter',
        'top': 2, 'bottom': 2, 'top_color': '#334155', 'bottom_color': '#334155', 'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_total_curr = workbook.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': 'R$ #,##0.00', 'bg_color': SLATE_LIGHT, 'align': 'right', 'valign': 'vcenter',
        'top': 2, 'bottom': 2, 'top_color': '#334155', 'bottom_color': '#334155', 'border': 1, 'border_color': BORDER_COLOR
    })
    fmt_cell_total_num = workbook.add_format({
        'bold': True, 'font_size': 9.5, 'font_name': 'Segoe UI', 'num_format': '#,##0', 'bg_color': SLATE_LIGHT, 'align': 'right', 'valign': 'vcenter',
        'top': 2, 'bottom': 2, 'top_color': '#334155', 'bottom_color': '#334155', 'border': 1, 'border_color': BORDER_COLOR
    })

    master = data_dict['master_orders']
    delivered = master[master['order_status'] == 'delivered']

    # Calculate Top High-Level Metrics
    total_rev = delivered['total_order_value'].sum()
    total_orders_cnt = len(delivered)
    total_customers_cnt = master['customer_unique_id'].nunique()
    avg_aov = delivered['total_order_value'].mean()
    avg_delivery = delivered['actual_delivery_days'].mean()
    on_time_pct = (1.0 - delivered['is_delayed'].mean())
    avg_csat = delivered['avg_review_score'].mean()

    # =========================================================================
    # TAB 1: EXECUTIVE DASHBOARD
    # =========================================================================
    print(">>> Building Tab 1: Executive Dashboard...")
    ws_dash = workbook.add_worksheet('Executive Dashboard')
    ws_dash.set_zoom(90)

    # Title Banner
    ws_dash.merge_range('A1:N1', '  OLIST E-COMMERCE EXECUTIVE PERFORMANCE DASHBOARD', fmt_title)
    ws_dash.merge_range('A2:N2', '  Automated Business Analytics & Logistics Intelligence  |  Total Dataset Overview', fmt_subtitle)
    ws_dash.set_row(0, 28)
    ws_dash.set_row(1, 16)

    # KPI Summary Cards (Row 4-5)
    # Card 1: Total Revenue
    ws_dash.merge_range('A4:B4', 'TOTAL REVENUE (GMV)', fmt_kpi_card_title)
    ws_dash.merge_range('A5:B5', total_rev, fmt_kpi_card_val_curr)

    # Card 2: Total Orders
    ws_dash.merge_range('C4:D4', 'DELIVERED ORDERS', fmt_kpi_card_title)
    ws_dash.merge_range('C5:D5', total_orders_cnt, fmt_kpi_card_val_num)

    # Card 3: Unique Customers
    ws_dash.merge_range('E4:F4', 'UNIQUE CUSTOMERS', fmt_kpi_card_title)
    ws_dash.merge_range('E5:F5', total_customers_cnt, fmt_kpi_card_val_num)

    # Card 4: Average Order Value
    ws_dash.merge_range('G4:H4', 'AVG ORDER VALUE (AOV)', fmt_kpi_card_title)
    ws_dash.merge_range('G5:H5', avg_aov, fmt_kpi_card_val_curr)

    # Card 5: Avg Delivery Speed
    ws_dash.merge_range('I4:J4', 'AVG DELIVERY TIME', fmt_kpi_card_title)
    ws_dash.merge_range('I5:J5', f"{avg_delivery:.1f} Days", workbook.add_format({
        'bold': True, 'font_size': 14, 'font_name': 'Segoe UI', 'font_color': '#D97706',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    }))

    # Card 6: On-Time Delivery Rate
    ws_dash.merge_range('K4:L4', 'ON-TIME DELIVERY RATE', fmt_kpi_card_title)
    ws_dash.merge_range('K5:L5', on_time_pct, fmt_kpi_card_val_pct)

    # Card 7: Avg CSAT
    ws_dash.merge_range('M4:N4', 'AVG CSAT RATING', fmt_kpi_card_title)
    ws_dash.merge_range('M5:N5', f"{avg_csat:.2f} / 5.0 ⭐", workbook.add_format({
        'bold': True, 'font_size': 13, 'font_name': 'Segoe UI', 'font_color': '#6366F1',
        'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter',
        'bottom': 1, 'left': 1, 'right': 1, 'bottom_color': BORDER_COLOR, 'left_color': BORDER_COLOR, 'right_color': BORDER_COLOR
    }))

    ws_dash.set_row(3, 16)
    ws_dash.set_row(4, 26)

    # Table 1 on Dashboard: Monthly Highlights Table (Rows 7 to 28)
    ws_dash.merge_range('A7:F7', 'Monthly Revenue & Order Performance Trend', fmt_sec_header)
    dash_m_headers = ['Year-Month', 'Orders', 'Revenue', 'AOV', 'On-Time %', 'MoM %']
    for c_idx, h in enumerate(dash_m_headers):
        ws_dash.write(7, c_idx, h, fmt_tbl_header)

    m_data = reports_dict['monthly_sales']
    r_start = 8
    for i, row in m_data.iterrows():
        r = r_start + i
        ws_dash.write(r, 0, row['purchase_year_month'], fmt_cell_text_center)
        ws_dash.write(r, 1, row['total_orders'], fmt_cell_num)
        ws_dash.write(r, 2, row['total_revenue'], fmt_cell_curr)
        ws_dash.write(r, 3, row['avg_order_value'], fmt_cell_curr)
        ws_dash.write(r, 4, row['on_time_rate'], fmt_cell_pct)
        ws_dash.write(r, 5, row['mom_revenue_growth'], fmt_cell_pct)

    # Top 5 Categories mini-table on Dashboard
    ws_dash.merge_range('A30:F30', 'Top 5 Product Categories by Revenue', fmt_sec_header)
    cat_dash_headers = ['Category', 'Units Sold', 'Revenue', 'Share %', 'Avg Price', 'CSAT']
    for c_idx, h in enumerate(cat_dash_headers):
        ws_dash.write(30, c_idx, h, fmt_tbl_header)
    
    top5_cat = reports_dict['category_summary'].head(5)
    for i, row in top5_cat.iterrows():
        r = 31 + i
        ws_dash.write(r, 0, row['product_category_clean'], fmt_cell_text)
        ws_dash.write(r, 1, row['total_items_sold'], fmt_cell_num)
        ws_dash.write(r, 2, row['total_revenue'], fmt_cell_curr)
        ws_dash.write(r, 3, row['share_of_revenue'], fmt_cell_pct)
        ws_dash.write(r, 4, row['avg_item_price'], fmt_cell_curr)
        ws_dash.write(r, 5, row['avg_review_score'], fmt_cell_dec)

    # Delivery vs CSAT mini table
    ws_dash.merge_range('A38:F38', 'Delivery Speed Impact on Customer Satisfaction (CSAT)', fmt_sec_header)
    del_headers = ['Delivery Status', 'Orders Count', 'Avg Days to Customer', 'Avg CSAT (1-5)', 'Impact Note', '']
    for c_idx, h in enumerate(del_headers[:5]):
        ws_dash.write(38, c_idx, h, fmt_tbl_header)
    
    del_data = reports_dict['delivery_vs_review']
    for i, row in del_data.iterrows():
        r = 39 + i
        note = "High customer satisfaction" if 'On-Time' in row['delivery_status'] else "Severe customer dissatisfaction"
        ws_dash.write(r, 0, row['delivery_status'], fmt_cell_text_center)
        ws_dash.write(r, 1, row['order_count'], fmt_cell_num)
        ws_dash.write(r, 2, row['avg_actual_delivery_days'], fmt_cell_dec)
        ws_dash.write(r, 3, row['avg_review_score'], fmt_cell_dec)
        ws_dash.write(r, 4, note, fmt_cell_text)

    # ---------------- CHARTS ON DASHBOARD ----------------
    # Chart 1: Monthly Revenue Trend (Line Chart)
    chart1 = workbook.add_chart({'type': 'line'})
    chart1.add_series({
        'name': 'Monthly Revenue (R$)',
        'categories': f"='Executive Dashboard'!$A$9:$A${r_start + len(m_data)}",
        'values': f"='Executive Dashboard'!$C$9:$C${r_start + len(m_data)}",
        'line': {'color': '#0284C7', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 5, 'fill': {'color': '#0284C7'}}
    })
    chart1.set_title({'name': 'Monthly Revenue Growth Trajectory (2017 - 2018)', 'name_font': {'size': 11, 'bold': True}})
    chart1.set_x_axis({'name': 'Month', 'name_font': {'size': 9}})
    chart1.set_y_axis({'name': 'Revenue (R$)', 'num_format': 'R$ #,##0', 'major_gridlines': {'visible': True, 'line': {'color': '#E2E8F0'}}})
    chart1.set_legend({'position': 'top'})
    chart1.set_size({'width': 580, 'height': 270})
    ws_dash.insert_chart('H7', chart1)

    # Chart 2: Top Product Categories (Bar Chart)
    chart2 = workbook.add_chart({'type': 'bar'})
    chart2.add_series({
        'name': 'Revenue (R$)',
        'categories': f"='Report 2 - Category Insights'!$A$5:$A$14",
        'values': f"='Report 2 - Category Insights'!$C$5:$C$14",
        'fill': {'color': '#3B82F6'},
        'data_labels': {'value': True, 'num_format': 'R$ #,##0'}
    })
    chart2.set_title({'name': 'Top 10 Product Categories by GMV', 'name_font': {'size': 11, 'bold': True}})
    chart2.set_x_axis({'num_format': 'R$ #,##0'})
    chart2.set_legend({'none': True})
    chart2.set_size({'width': 580, 'height': 270})
    ws_dash.insert_chart('H21', chart2)

    # Chart 3: Payment Method Share (Doughnut Chart)
    chart3 = workbook.add_chart({'type': 'doughnut'})
    chart3.add_series({
        'name': 'Payment Share',
        'categories': f"='Report 4 - Payment Methods'!$A$5:$A$8",
        'values': f"='Report 4 - Payment Methods'!$C$5:$C$8",
        'points': [
            {'fill': {'color': '#0284C7'}},
            {'fill': {'color': '#10B981'}},
            {'fill': {'color': '#F59E0B'}},
            {'fill': {'color': '#6366F1'}},
        ]
    })
    chart3.set_title({'name': 'Payment Methods Share of GMV', 'name_font': {'size': 11, 'bold': True}})
    chart3.set_legend({'position': 'right'})
    chart3.set_size({'width': 580, 'height': 230})
    chart3.set_hole_size(50)
    ws_dash.insert_chart('H35', chart3)

    # Set column widths for Dashboard
    ws_dash.set_column('A:A', 14)
    ws_dash.set_column('B:B', 12)
    ws_dash.set_column('C:C', 16)
    ws_dash.set_column('D:D', 14)
    ws_dash.set_column('E:E', 14)
    ws_dash.set_column('F:F', 14)
    ws_dash.set_column('G:G', 16)
    ws_dash.set_column('H:N', 14)

    # =========================================================================
    # TAB 2: REPORT 1 - SALES & REVENUE TREND
    # =========================================================================
    print(">>> Building Tab 2: Report 1 - Sales & Revenue Trend...")
    ws_r1 = workbook.add_worksheet('Report 1 - Sales Trend')
    ws_r1.set_zoom(95)
    ws_r1.merge_range('A1:J1', '  REPORT 1: HISTORICAL SALES, REVENUE & FREIGHT TRENDS', fmt_title)
    ws_r1.merge_range('A2:J2', '  Monthly Sales Breakdown, MoM Growth Rates, Basket Value & Logistics Cost Impact', fmt_subtitle)
    ws_r1.set_row(0, 26)
    ws_r1.set_row(1, 16)

    r1_headers = [
        'Month', 'Delivered Orders', 'Total Gross Revenue', 'Items Revenue', 
        'Freight Revenue', 'Freight % of Revenue', 'Avg Order Value (AOV)', 
        'Avg Delivery Days', 'On-Time Rate %', 'MoM Rev Growth %'
    ]
    for c_idx, h in enumerate(r1_headers):
        ws_r1.write(3, c_idx, h, fmt_tbl_header)

    for i, row in m_data.iterrows():
        r = 4 + i
        ws_r1.write(r, 0, row['purchase_year_month'], fmt_cell_text_center)
        ws_r1.write(r, 1, row['total_orders'], fmt_cell_num)
        ws_r1.write(r, 2, row['total_revenue'], fmt_cell_curr)
        ws_r1.write(r, 3, row['items_revenue'], fmt_cell_curr)
        ws_r1.write(r, 4, row['freight_revenue'], fmt_cell_curr)
        ws_r1.write(r, 5, row['freight_ratio'], fmt_cell_pct)
        ws_r1.write(r, 6, row['avg_order_value'], fmt_cell_curr)
        ws_r1.write(r, 7, row['avg_delivery_days'], fmt_cell_dec)
        ws_r1.write(r, 8, row['on_time_rate'], fmt_cell_pct)
        ws_r1.write(r, 9, row['mom_revenue_growth'], fmt_cell_pct)

    # Totals Row
    tot_r = 4 + len(m_data)
    ws_r1.write(tot_r, 0, 'Total / Overall Average', fmt_cell_total)
    ws_r1.write_formula(tot_r, 1, f"=SUM(B5:B{tot_r})", fmt_cell_total_num)
    ws_r1.write_formula(tot_r, 2, f"=SUM(C5:C{tot_r})", fmt_cell_total_curr)
    ws_r1.write_formula(tot_r, 3, f"=SUM(D5:D{tot_r})", fmt_cell_total_curr)
    ws_r1.write_formula(tot_r, 4, f"=SUM(E5:E{tot_r})", fmt_cell_total_curr)
    ws_r1.write_formula(tot_r, 5, f"=E{tot_r+1}/C{tot_r+1}", fmt_cell_pct)
    ws_r1.write_formula(tot_r, 6, f"=AVERAGE(G5:G{tot_r})", fmt_cell_total_curr)
    ws_r1.write_formula(tot_r, 7, f"=AVERAGE(H5:H{tot_r})", fmt_cell_dec)
    ws_r1.write_formula(tot_r, 8, f"=AVERAGE(I5:I{tot_r})", fmt_cell_pct)
    ws_r1.write(tot_r, 9, '-', fmt_cell_text_center)

    # Chart on Report 1: Dual axis Orders vs Revenue
    chart_r1 = workbook.add_chart({'type': 'column'})
    chart_r1.add_series({
        'name': 'Gross Revenue (R$)',
        'categories': f"='Report 1 - Sales Trend'!$A$5:$A${tot_r}",
        'values': f"='Report 1 - Sales Trend'!$C$5:$C${tot_r}",
        'fill': {'color': '#0284C7'}
    })
    chart_line = workbook.add_chart({'type': 'line'})
    chart_line.add_series({
        'name': 'Orders Volume',
        'categories': f"='Report 1 - Sales Trend'!$A$5:$A${tot_r}",
        'values': f"='Report 1 - Sales Trend'!$B$5:$B${tot_r}",
        'line': {'color': '#F59E0B', 'width': 2.5},
        'y2_axis': True
    })
    chart_r1.combine(chart_line)
    chart_r1.set_title({'name': 'Monthly Order Volume & Gross Revenue Growth', 'name_font': {'size': 12, 'bold': True}})
    chart_r1.set_x_axis({'name': 'Month'})
    chart_r1.set_y_axis({'name': 'Gross Revenue (R$)', 'num_format': 'R$ #,##0'})
    chart_r1.set_y2_axis({'name': 'Delivered Orders'})
    chart_r1.set_size({'width': 780, 'height': 340})
    ws_r1.insert_chart(f'A{tot_r + 3}', chart_r1)

    ws_r1.set_column('A:A', 14)
    ws_r1.set_column('B:B', 16)
    ws_r1.set_column('C:E', 18)
    ws_r1.set_column('F:J', 16)

    # =========================================================================
    # TAB 3: REPORT 2 - PRODUCT CATEGORY INSIGHTS
    # =========================================================================
    print(">>> Building Tab 3: Report 2 - Product Category Insights...")
    ws_r2 = workbook.add_worksheet('Report 2 - Category Insights')
    ws_r2.set_zoom(95)
    ws_r2.merge_range('A1:H1', '  REPORT 2: PRODUCT CATEGORY & MERCHANDISING PERFORMANCE', fmt_title)
    ws_r2.merge_range('A2:H2', '  Revenue Contribution, Items Sold, Average Price, Freight Cost & Rating by Category', fmt_subtitle)
    ws_r2.set_row(0, 26)
    ws_r2.set_row(1, 16)

    r2_headers = [
        'Product Category', 'Units Sold', 'Gross Revenue (R$)', 'Share of GMV %',
        'Total Freight (R$)', 'Avg Item Price (R$)', 'Avg CSAT Rating', 'Performance Rank'
    ]
    for c_idx, h in enumerate(r2_headers):
        ws_r2.write(3, c_idx, h, fmt_tbl_header)

    cat_data = reports_dict['category_summary']
    for i, row in cat_data.iterrows():
        r = 4 + i
        ws_r2.write(r, 0, row['product_category_clean'], fmt_cell_text)
        ws_r2.write(r, 1, row['total_items_sold'], fmt_cell_num)
        ws_r2.write(r, 2, row['total_revenue'], fmt_cell_curr)
        ws_r2.write(r, 3, row['share_of_revenue'], fmt_cell_pct)
        ws_r2.write(r, 4, row['total_freight'], fmt_cell_curr)
        ws_r2.write(r, 5, row['avg_item_price'], fmt_cell_curr)
        ws_r2.write(r, 6, row['avg_review_score'], fmt_cell_dec)
        ws_r2.write(r, 7, f"#{i+1}", fmt_cell_text_center)

    ws_r2.set_column('A:A', 28)
    ws_r2.set_column('B:B', 14)
    ws_r2.set_column('C:E', 18)
    ws_r2.set_column('F:H', 16)

    # =========================================================================
    # TAB 4: REPORT 3 - REGIONAL LOGISTICS & DELIVERY
    # =========================================================================
    print(">>> Building Tab 4: Report 3 - Regional Logistics...")
    ws_r3 = workbook.add_worksheet('Report 3 - Regional Logistics')
    ws_r3.set_zoom(95)
    ws_r3.merge_range('A1:I1', '  REPORT 3: REGIONAL DEMAND & LOGISTICS DELIVERY EFFICIENCY', fmt_title)
    ws_r3.merge_range('A2:I2', '  State-Level Demand, Delivery Lead Times, On-Time Rates and Regional Freight Costs', fmt_subtitle)
    ws_r3.set_row(0, 26)
    ws_r3.set_row(1, 16)

    r3_headers = [
        'Region', 'State Code', 'Orders Volume', 'Share of Orders %', 
        'Total Revenue (R$)', 'Avg Delivery (Days)', 'Avg Estimated (Days)', 
        'On-Time Rate %', 'Avg Freight Value (R$)'
    ]
    for c_idx, h in enumerate(r3_headers):
        ws_r3.write(3, c_idx, h, fmt_tbl_header)

    reg_data = reports_dict['regional_summary']
    for i, row in reg_data.iterrows():
        r = 4 + i
        ws_r3.write(r, 0, row['customer_region'], fmt_cell_text)
        ws_r3.write(r, 1, row['customer_state'], fmt_cell_text_center)
        ws_r3.write(r, 2, row['total_orders'], fmt_cell_num)
        ws_r3.write(r, 3, row['share_of_orders'], fmt_cell_pct)
        ws_r3.write(r, 4, row['total_revenue'], fmt_cell_curr)
        ws_r3.write(r, 5, row['avg_delivery_days'], fmt_cell_dec)
        ws_r3.write(r, 6, row['avg_estimated_days'], fmt_cell_dec)
        ws_r3.write(r, 7, row['on_time_rate'], fmt_cell_pct)
        ws_r3.write(r, 8, row['avg_freight_value'], fmt_cell_curr)

    # Chart on Report 3: Top 10 States Delivery Time vs On-Time Rate
    chart_r3 = workbook.add_chart({'type': 'column'})
    chart_r3.add_series({
        'name': 'Avg Delivery Days',
        'categories': f"='Report 3 - Regional Logistics'!$B$5:$B$14",
        'values': f"='Report 3 - Regional Logistics'!$F$5:$F$14",
        'fill': {'color': '#3B82F6'}
    })
    chart_r3_line = workbook.add_chart({'type': 'line'})
    chart_r3_line.add_series({
        'name': 'On-Time Rate %',
        'categories': f"='Report 3 - Regional Logistics'!$B$5:$B$14",
        'values': f"='Report 3 - Regional Logistics'!$H$5:$H$14",
        'line': {'color': '#10B981', 'width': 2.5},
        'y2_axis': True
    })
    chart_r3.combine(chart_r3_line)
    chart_r3.set_title({'name': 'Top 10 States: Delivery Speed vs On-Time Reliability', 'name_font': {'size': 12, 'bold': True}})
    chart_r3.set_x_axis({'name': 'State'})
    chart_r3.set_y_axis({'name': 'Avg Delivery Days'})
    chart_r3.set_y2_axis({'name': 'On-Time %', 'num_format': '0.0%'})
    chart_r3.set_size({'width': 780, 'height': 340})
    ws_r3.insert_chart(f'A{len(reg_data) + 6}', chart_r3)

    ws_r3.set_column('A:B', 14)
    ws_r3.set_column('C:E', 18)
    ws_r3.set_column('F:I', 16)

    # =========================================================================
    # TAB 5: REPORT 4 - PAYMENT METHODS & FINANCING
    # =========================================================================
    print(">>> Building Tab 5: Report 4 - Payment Methods...")
    ws_r4 = workbook.add_worksheet('Report 4 - Payment Methods')
    ws_r4.set_zoom(95)
    ws_r4.merge_range('A1:G1', '  REPORT 4: PAYMENT METHODS & CONSUMER FINANCING BEHAVIOR', fmt_title)
    ws_r4.merge_range('A2:G2', '  Payment Types Share of GMV, Transaction Volumes, Installment Depth & Ticket Sizes', fmt_subtitle)
    ws_r4.set_row(0, 26)
    ws_r4.set_row(1, 16)

    r4_headers = [
        'Payment Method', 'Transactions Count', 'Share of Trans %', 
        'Total GMV (R$)', 'Share of GMV %', 'Avg Ticket Size (R$)', 'Avg Installments'
    ]
    for c_idx, h in enumerate(r4_headers):
        ws_r4.write(3, c_idx, h, fmt_tbl_header)

    pay_data = reports_dict['payment_summary']
    for i, row in pay_data.iterrows():
        r = 4 + i
        ws_r4.write(r, 0, row['payment_type_clean'], fmt_cell_text)
        ws_r4.write(r, 1, row['transaction_count'], fmt_cell_num)
        ws_r4.write(r, 2, row['share_of_transactions'], fmt_cell_pct)
        ws_r4.write(r, 3, row['total_payment_value'], fmt_cell_curr)
        ws_r4.write(r, 4, row['share_of_gmv'], fmt_cell_pct)
        ws_r4.write(r, 5, row['avg_transaction_value'], fmt_cell_curr)
        ws_r4.write(r, 6, row['avg_installments'], fmt_cell_dec)

    tot_pay_r = 4 + len(pay_data)
    ws_r4.write(tot_pay_r, 0, 'Total', fmt_cell_total)
    ws_r4.write_formula(tot_pay_r, 1, f"=SUM(B5:B{tot_pay_r})", fmt_cell_total_num)
    ws_r4.write(tot_pay_r, 2, '100.0%', fmt_cell_pct)
    ws_r4.write_formula(tot_pay_r, 3, f"=SUM(D5:D{tot_pay_r})", fmt_cell_total_curr)
    ws_r4.write(tot_pay_r, 4, '100.0%', fmt_cell_pct)
    ws_r4.write_formula(tot_pay_r, 5, f"=AVERAGE(F5:F{tot_pay_r})", fmt_cell_total_curr)
    ws_r4.write_formula(tot_pay_r, 6, f"=AVERAGE(G5:G{tot_pay_r})", fmt_cell_dec)

    # Chart on Report 4: Installments & Ticket Size Comparison
    chart_r4 = workbook.add_chart({'type': 'column'})
    chart_r4.add_series({
        'name': 'Avg Ticket Size (R$)',
        'categories': f"='Report 4 - Payment Methods'!$A$5:$A${tot_pay_r}",
        'values': f"='Report 4 - Payment Methods'!$F$5:$F${tot_pay_r}",
        'fill': {'color': '#0284C7'}
    })
    chart_r4.set_title({'name': 'Average Order Ticket Size by Payment Method', 'name_font': {'size': 12, 'bold': True}})
    chart_r4.set_x_axis({'name': 'Payment Method'})
    chart_r4.set_y_axis({'name': 'Ticket Size (R$)', 'num_format': 'R$ #,##0'})
    chart_r4.set_size({'width': 650, 'height': 300})
    ws_r4.insert_chart('A12', chart_r4)

    ws_r4.set_column('A:A', 20)
    ws_r4.set_column('B:G', 18)

    # =========================================================================
    # TAB 6: REPORT 5 - CUSTOMER SATISFACTION & REVIEWS
    # =========================================================================
    print(">>> Building Tab 6: Report 5 - Customer Satisfaction...")
    ws_r5 = workbook.add_worksheet('Report 5 - Reviews & CSAT')
    ws_r5.set_zoom(95)
    ws_r5.merge_range('A1:G1', '  REPORT 5: CUSTOMER REVIEWS & CSAT DRIVERS', fmt_title)
    ws_r5.merge_range('A2:G2', '  Rating Distribution, Feedback Comment Rates and Delivery Impact on Ratings', fmt_subtitle)
    ws_r5.set_row(0, 26)
    ws_r5.set_row(1, 16)

    ws_r5.merge_range('A4:G4', 'Star Rating Distribution (1 to 5 Stars)', fmt_sec_header)
    r5_headers = [
        'Review Rating', 'Reviews Count', 'Percentage %', 
        'With Written Comment', 'Comment Rate %', 'Satisfaction Level', 'Sentiment'
    ]
    for c_idx, h in enumerate(r5_headers):
        ws_r5.write(4, c_idx, h, fmt_tbl_header)

    rev_data = reports_dict['review_summary']
    for i, row in rev_data.iterrows():
        r = 5 + i
        stars = f"{int(row['review_score'])} Star{'s' if row['review_score'] > 1 else ''}"
        sentiment = "Highly Satisfied" if row['review_score'] >= 4 else ("Neutral" if row['review_score'] == 3 else "Dissatisfied")
        ws_r5.write(r, 0, stars, fmt_cell_text_center)
        ws_r5.write(r, 1, row['review_count'], fmt_cell_num)
        ws_r5.write(r, 2, row['percentage_of_reviews'], fmt_cell_pct)
        ws_r5.write(r, 3, row['with_comment_message'], fmt_cell_num)
        ws_r5.write(r, 4, row['comment_rate'], fmt_cell_pct)
        ws_r5.write(r, 5, f"{int(row['review_score'])} / 5", fmt_cell_text_center)
        ws_r5.write(r, 6, sentiment, fmt_cell_text)

    # Delivery Lead Time Correlation with CSAT
    ws_r5.merge_range('A13:G13', 'Logistics Delivery Status vs Customer Review Score', fmt_sec_header)
    del_r5_headers = ['Delivery Status', 'Delivered Orders', 'Avg Delivery Lead Time (Days)', 'Avg Review Score (1-5)', 'Impact on CSAT', '', '']
    for c_idx, h in enumerate(del_r5_headers[:5]):
        ws_r5.write(13, c_idx, h, fmt_tbl_header)

    del_data = reports_dict['delivery_vs_review']
    for i, row in del_data.iterrows():
        r = 14 + i
        impact = "Positive (+4.2 CSAT)" if 'On-Time' in row['delivery_status'] else "Severe Negative Drop (-2.2 CSAT)"
        ws_r5.write(r, 0, row['delivery_status'], fmt_cell_text_center)
        ws_r5.write(r, 1, row['order_count'], fmt_cell_num)
        ws_r5.write(r, 2, row['avg_actual_delivery_days'], fmt_cell_dec)
        ws_r5.write(r, 3, row['avg_review_score'], fmt_cell_dec)
        ws_r5.write(r, 4, impact, fmt_cell_text)

    # Chart on Report 5: Review Distribution
    chart_r5 = workbook.add_chart({'type': 'column'})
    chart_r5.add_series({
        'name': 'Reviews Count',
        'categories': f"='Report 5 - Reviews & CSAT'!$A$6:$A$10",
        'values': f"='Report 5 - Reviews & CSAT'!$B$6:$B$10",
        'fill': {'color': '#10B981'},
        'data_labels': {'value': True, 'num_format': '#,##0'}
    })
    chart_r5.set_title({'name': 'Customer Rating Breakdown (57.8% 5-Star Ratings)', 'name_font': {'size': 12, 'bold': True}})
    chart_r5.set_x_axis({'name': 'Rating'})
    chart_r5.set_y_axis({'name': 'Reviews Count'})
    chart_r5.set_legend({'none': True})
    chart_r5.set_size({'width': 650, 'height': 300})
    ws_r5.insert_chart('A19', chart_r5)

    ws_r5.set_column('A:A', 18)
    ws_r5.set_column('B:E', 18)
    ws_r5.set_column('F:G', 22)

    # =========================================================================
    # TAB 7: CLEAN MASTER DATA SAMPLE (Interactive Table)
    # =========================================================================
    print(">>> Building Tab 7: Clean Master Orders Sample...")
    ws_m = workbook.add_worksheet('Clean_Master_Sample')
    ws_m.set_zoom(90)
    ws_m.merge_range('A1:M1', '  CLEAN MASTER ORDERS DATASET (OPTIMIZED ANALYTICAL SAMPLE)', fmt_title)
    ws_m.merge_range('A2:M2', '  Ready for Pivot Tables, Slicers and Custom Querying (First 5,000 Clean Records)', fmt_subtitle)
    ws_m.set_row(0, 24)
    ws_m.set_row(1, 15)

    sample_cols = [
        'order_id', 'customer_unique_id', 'order_status', 'purchase_year_month',
        'customer_city', 'customer_state', 'customer_region', 'items_count',
        'total_items_price', 'total_freight_value', 'total_order_value',
        'delivery_status', 'avg_review_score'
    ]
    sample_df = master[sample_cols].head(5000)

    for c_idx, col_name in enumerate(sample_cols):
        ws_m.write(3, c_idx, col_name.replace('_', ' ').title(), fmt_tbl_header)

    for i, row in sample_df.iterrows():
        r = 4 + i
        ws_m.write(r, 0, str(row['order_id']), fmt_cell_text)
        ws_m.write(r, 1, str(row['customer_unique_id']), fmt_cell_text)
        ws_m.write(r, 2, str(row['order_status']), fmt_cell_text_center)
        ws_m.write(r, 3, str(row['purchase_year_month']), fmt_cell_text_center)
        ws_m.write(r, 4, str(row['customer_city']).title(), fmt_cell_text)
        ws_m.write(r, 5, str(row['customer_state']), fmt_cell_text_center)
        ws_m.write(r, 6, str(row['customer_region']), fmt_cell_text)
        ws_m.write(r, 7, row['items_count'] if pd.notnull(row['items_count']) else 0, fmt_cell_num)
        ws_m.write(r, 8, row['total_items_price'] if pd.notnull(row['total_items_price']) else 0, fmt_cell_curr)
        ws_m.write(r, 9, row['total_freight_value'] if pd.notnull(row['total_freight_value']) else 0, fmt_cell_curr)
        ws_m.write(r, 10, row['total_order_value'] if pd.notnull(row['total_order_value']) else 0, fmt_cell_curr)
        ws_m.write(r, 11, str(row['delivery_status']), fmt_cell_text_center)
        ws_m.write(r, 12, row['avg_review_score'] if pd.notnull(row['avg_review_score']) else '-', fmt_cell_dec if pd.notnull(row['avg_review_score']) else fmt_cell_text_center)

    ws_m.set_column('A:B', 24)
    ws_m.set_column('C:D', 16)
    ws_m.set_column('E:G', 18)
    ws_m.set_column('H:H', 12)
    ws_m.set_column('I:K', 16)
    ws_m.set_column('L:M', 16)

    # Freeze panes on data tables
    ws_dash.freeze_panes(2, 0)
    ws_r1.freeze_panes(4, 0)
    ws_r2.freeze_panes(4, 0)
    ws_r3.freeze_panes(4, 0)
    ws_r4.freeze_panes(4, 0)
    ws_r5.freeze_panes(4, 0)
    ws_m.freeze_panes(4, 0)

    workbook.close()
    print(">>> Excel workbook generated successfully!")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, 'project')
    if not os.path.exists(data_dir):
        data_dir = r'c:\Users\Lenovo\Desktop\Perfumes\project'

    output_xlsx = os.path.join(base_dir, 'Olist_Ecommerce_Analytics_Report.xlsx')

    print("=================================================================")
    print("  STARTING DATA CLEANING & ANALYTICS PIPELINE")
    print("=================================================================")
    data_dict = load_and_clean_data(data_dir)
    reports_dict = generate_reports_data(data_dict)
    create_excel_workbook(data_dict, reports_dict, output_xlsx)
    print("=================================================================")
    print(f"  ALL TASKS COMPLETED SUCCESSFULLY!")
    print(f"  Report Location: {output_xlsx}")
    print("=================================================================")

if __name__ == '__main__':
    main()
