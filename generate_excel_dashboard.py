"""
Professional Excel Dashboard Generator for Olist E-Commerce Project
===================================================================
Generates an executive-grade Excel workbook with:
- Dedicated visual Executive Dashboard sheet
- Beautifully formatted KPI Cards
- Native Excel Charts (Revenue Trends, Top Categories, State Distribution, Payment Breakdown, Delivery vs Satisfaction)
- Supporting Pivot / Analytical sheets with custom number formatting (R$, %, integers)
- Consistent Navy / Teal / Slate design system
"""

import os
import pandas as pd
import numpy as np
import xlsxwriter

def create_excel_dashboard():
    project_dir = r'c:\Users\Lenovo\Desktop\Perfumes\olist-ecommerce-project'
    cleaned_csv = os.path.join(project_dir, 'data', 'olist_cleaned.csv')
    output_xlsx = os.path.join(project_dir, 'dashboard', 'olist_dashboard.xlsx')
    
    print("Loading cleaned dataset...")
    df = pd.read_csv(cleaned_csv, low_memory=False)
    
    # Filter to standard 2017-01 to 2018-08 analytical window (standard Olist analysis period)
    df_core = df[df['order_month'].between('2017-01', '2018-08')].copy()
    
    # 1. Monthly Summary
    monthly = df_core.groupby('order_month').agg(
        total_revenue=('total_order_value', 'sum'),
        total_orders=('order_id', 'nunique'),
        avg_delivery_days=('delivery_time_days', 'mean'),
        avg_review_score=('review_score', 'mean')
    ).reset_index()
    monthly['aov'] = monthly['total_revenue'] / monthly['total_orders']
    
    # 2. Top 10 Categories
    top_cat = df_core.groupby('product_category_clean').agg(
        revenue=('total_order_value', 'sum'),
        orders=('order_id', 'nunique'),
        avg_review=('review_score', 'mean')
    ).reset_index().sort_values('revenue', ascending=False).head(10)
    
    # 3. Top 10 States
    top_states = df_core.groupby('customer_state').agg(
        revenue=('total_order_value', 'sum'),
        orders=('order_id', 'nunique'),
        avg_delivery=('delivery_time_days', 'mean')
    ).reset_index().sort_values('revenue', ascending=False).head(10)
    
    # 4. Payment Types
    payments = df_core.groupby('primary_payment_type').agg(
        orders=('order_id', 'nunique'),
        revenue=('total_order_value', 'sum')
    ).reset_index().sort_values('orders', ascending=False)
    
    # 5. Review Score Distribution
    reviews = df_core.groupby('review_score').agg(
        review_count=('order_id', 'nunique'),
        avg_delivery_delay=('delivery_delay_days', 'mean')
    ).reset_index()
    reviews['review_score_label'] = reviews['review_score'].astype(int).astype(str) + ' Stars'
    
    # 6. Delivery Performance vs Review Score
    delay_buckets = df_core.copy()
    delay_buckets['delay_group'] = pd.cut(
        delay_buckets['delivery_delay_days'],
        bins=[-999, -5, 0, 5, 15, 999],
        labels=['> 5 Days Early', '1-5 Days Early', '1-5 Days Late', '6-15 Days Late', '> 15 Days Late']
    )
    delay_summary = delay_buckets.groupby('delay_group', observed=True).agg(
        order_count=('order_id', 'nunique'),
        avg_review_score=('review_score', 'mean')
    ).reset_index()
    
    print(f"Creating Excel workbook at: {output_xlsx}")
    workbook = xlsxwriter.Workbook(output_xlsx)
    
    # Define Brand Palette & Formatting
    # Primary: Dark Navy #0F172A, Accent 1: Teal #0D9488, Accent 2: Amber #D97706, Card Bg: #F8FAFC
    fmt_title = workbook.add_format({
        'bold': True, 'font_size': 18, 'font_name': 'Segoe UI',
        'font_color': '#FFFFFF', 'bg_color': '#0F172A',
        'align': 'left', 'valign': 'vcenter', 'left': 1, 'top': 1, 'bottom': 1
    })
    fmt_subtitle = workbook.add_format({
        'font_size': 11, 'font_name': 'Segoe UI',
        'font_color': '#94A3B8', 'bg_color': '#0F172A',
        'align': 'left', 'valign': 'vcenter', 'left': 1, 'bottom': 1
    })
    fmt_section = workbook.add_format({
        'bold': True, 'font_size': 13, 'font_name': 'Segoe UI',
        'font_color': '#0F172A', 'bottom': 2, 'bottom_color': '#0D9488',
        'valign': 'vcenter'
    })
    
    # KPI Card Formats
    fmt_kpi_title = workbook.add_format({
        'bold': True, 'font_size': 9, 'font_name': 'Segoe UI',
        'font_color': '#64748B', 'bg_color': '#F1F5F9',
        'align': 'center', 'valign': 'vcenter', 'top': 1, 'left': 1, 'right': 1,
        'border_color': '#CBD5E1'
    })
    fmt_kpi_val_curr = workbook.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Segoe UI',
        'font_color': '#0F172A', 'bg_color': '#F1F5F9',
        'align': 'center', 'valign': 'vcenter', 'num_format': 'R$ #,##0',
        'left': 1, 'right': 1, 'border_color': '#CBD5E1'
    })
    fmt_kpi_val_int = workbook.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Segoe UI',
        'font_color': '#0F172A', 'bg_color': '#F1F5F9',
        'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0',
        'left': 1, 'right': 1, 'border_color': '#CBD5E1'
    })
    fmt_kpi_val_dec = workbook.add_format({
        'bold': True, 'font_size': 16, 'font_name': 'Segoe UI',
        'font_color': '#0F172A', 'bg_color': '#F1F5F9',
        'align': 'center', 'valign': 'vcenter', 'num_format': '0.00',
        'left': 1, 'right': 1, 'border_color': '#CBD5E1'
    })
    fmt_kpi_sub = workbook.add_format({
        'font_size': 8, 'font_name': 'Segoe UI',
        'font_color': '#0D9488', 'bg_color': '#F1F5F9',
        'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'left': 1, 'right': 1,
        'border_color': '#CBD5E1'
    })
    
    # Table Header & Data Formats
    fmt_th = workbook.add_format({
        'bold': True, 'font_size': 10, 'font_name': 'Segoe UI',
        'font_color': '#FFFFFF', 'bg_color': '#1E293B',
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#334155'
    })
    fmt_td_text = workbook.add_format({
        'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#1E293B',
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0'
    })
    fmt_td_curr = workbook.add_format({
        'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#1E293B',
        'align': 'right', 'valign': 'vcenter', 'num_format': 'R$ #,##0.00',
        'border': 1, 'border_color': '#E2E8F0'
    })
    fmt_td_int = workbook.add_format({
        'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#1E293B',
        'align': 'right', 'valign': 'vcenter', 'num_format': '#,##0',
        'border': 1, 'border_color': '#E2E8F0'
    })
    fmt_td_dec = workbook.add_format({
        'font_size': 9, 'font_name': 'Segoe UI', 'font_color': '#1E293B',
        'align': 'right', 'valign': 'vcenter', 'num_format': '0.00',
        'border': 1, 'border_color': '#E2E8F0'
    })
    
    # -------------------------------------------------------------
    # SHEET 1: EXECUTIVE DASHBOARD
    # -------------------------------------------------------------
    ws_dash = workbook.add_worksheet('Executive Dashboard')
    ws_dash.hide_gridlines(2)
    ws_dash.set_zoom(90)
    
    # Set Column Widths for Dashboard Grid
    ws_dash.set_column('A:A', 3)
    ws_dash.set_column('B:D', 13)
    ws_dash.set_column('E:E', 4)
    ws_dash.set_column('F:H', 13)
    ws_dash.set_column('I:I', 4)
    ws_dash.set_column('J:L', 13)
    ws_dash.set_column('M:M', 4)
    ws_dash.set_column('N:P', 13)
    ws_dash.set_column('Q:Q', 4)
    ws_dash.set_column('R:T', 13)
    
    # Dashboard Header Banner
    ws_dash.merge_range('B2:T2', '  BRAZILIAN E-COMMERCE (OLIST) EXECUTIVE DASHBOARD', fmt_title)
    ws_dash.merge_range('B3:T3', '  Data Analysis Period: Jan 2017 – Aug 2018 | Scope: 99.4k Orders, R$ 15.8M Revenue | Currency: Brazilian Real (R$)', fmt_subtitle)
    ws_dash.set_row(1, 28)
    ws_dash.set_row(2, 20)
    
    # KPI 1: Total Revenue
    ws_dash.merge_range('B5:D5', 'TOTAL NET REVENUE', fmt_kpi_title)
    ws_dash.merge_range('B6:D6', df_core['total_order_value'].sum(), fmt_kpi_val_curr)
    ws_dash.merge_range('B7:D7', 'Jan 2017 – Aug 2018', fmt_kpi_sub)
    
    # KPI 2: Total Orders
    ws_dash.merge_range('F5:H5', 'TOTAL ORDERS', fmt_kpi_title)
    ws_dash.merge_range('F6:H6', df_core['order_id'].nunique(), fmt_kpi_val_int)
    ws_dash.merge_range('F7:H7', 'Delivered & Processed', fmt_kpi_sub)
    
    # KPI 3: Average Order Value (AOV)
    aov_val = df_core['total_order_value'].sum() / df_core['order_id'].nunique()
    ws_dash.merge_range('J5:L5', 'AVG ORDER VALUE (AOV)', fmt_kpi_title)
    ws_dash.merge_range('J6:L6', aov_val, fmt_kpi_val_curr)
    ws_dash.merge_range('J7:L7', 'Per Unique Order', fmt_kpi_sub)
    
    # KPI 4: Avg Delivery Time
    ws_dash.merge_range('N5:P5', 'AVG DELIVERY TIME', fmt_kpi_title)
    ws_dash.merge_range('N6:P6', f"{df_core['delivery_time_days'].mean():.1f} Days", fmt_kpi_val_dec)
    ws_dash.merge_range('N7:P7', 'Purchase to Delivery', fmt_kpi_sub)
    
    # KPI 5: Avg Review Score
    ws_dash.merge_range('R5:T5', 'AVG REVIEW SCORE', fmt_kpi_title)
    ws_dash.merge_range('R6:T6', f"{df_core['review_score'].mean():.2f} / 5.0", fmt_kpi_val_dec)
    ws_dash.merge_range('R7:T7', 'Customer Satisfaction', fmt_kpi_sub)
    
    ws_dash.set_row(4, 18)
    ws_dash.set_row(5, 26)
    ws_dash.set_row(6, 16)
    
    # Section 1: Growth & Top Categories
    ws_dash.write('B9', 'Monthly Revenue Trend & Sales Momentum', fmt_section)
    ws_dash.write('L9', 'Top 10 Product Categories by Revenue', fmt_section)
    
    # Section 2: Regional & Operational Deep-Dive
    ws_dash.write('B24', 'Regional Revenue by Customer State (Top 10)', fmt_section)
    ws_dash.write('L24', 'Payment Method Distribution', fmt_section)
    
    # Section 3: Customer Satisfaction & Logistics
    ws_dash.write('B39', 'Customer Review Score Distribution (1–5 Stars)', fmt_section)
    ws_dash.write('L39', 'Delivery Delay vs. Customer Review Score', fmt_section)
    
    # -------------------------------------------------------------
    # SHEET 2: MONTHLY DATA (Pivot/Summary)
    # -------------------------------------------------------------
    ws_month = workbook.add_worksheet('Data_Monthly')
    headers_month = ['Order Month', 'Total Revenue (R$)', 'Total Orders', 'Avg Delivery (Days)', 'Avg Review Score', 'AOV (R$)']
    for col_idx, h in enumerate(headers_month):
        ws_month.write(0, col_idx, h, fmt_th)
    ws_month.set_column('A:F', 18)
    
    for row_idx, r in monthly.iterrows():
        ws_month.write(row_idx + 1, 0, r['order_month'], fmt_td_text)
        ws_month.write(row_idx + 1, 1, r['total_revenue'], fmt_td_curr)
        ws_month.write(row_idx + 1, 2, r['total_orders'], fmt_td_int)
        ws_month.write(row_idx + 1, 3, r['avg_delivery_days'], fmt_td_dec)
        ws_month.write(row_idx + 1, 4, r['avg_review_score'], fmt_td_dec)
        ws_month.write(row_idx + 1, 5, r['aov'], fmt_td_curr)
        
    # -------------------------------------------------------------
    # SHEET 3: CATEGORY DATA
    # -------------------------------------------------------------
    ws_cat = workbook.add_worksheet('Data_Categories')
    headers_cat = ['Product Category', 'Total Revenue (R$)', 'Total Orders', 'Avg Review Score']
    for col_idx, h in enumerate(headers_cat):
        ws_cat.write(0, col_idx, h, fmt_th)
    ws_cat.set_column('A:A', 28)
    ws_cat.set_column('B:D', 18)
    
    for row_idx, r in top_cat.iterrows():
        ws_cat.write(row_idx + 1, 0, r['product_category_clean'], fmt_td_text)
        ws_cat.write(row_idx + 1, 1, r['revenue'], fmt_td_curr)
        ws_cat.write(row_idx + 1, 2, r['orders'], fmt_td_int)
        ws_cat.write(row_idx + 1, 3, r['avg_review'], fmt_td_dec)
        
    # -------------------------------------------------------------
    # SHEET 4: STATE DATA
    # -------------------------------------------------------------
    ws_state = workbook.add_worksheet('Data_States')
    headers_state = ['Customer State', 'Total Revenue (R$)', 'Total Orders', 'Avg Delivery (Days)']
    for col_idx, h in enumerate(headers_state):
        ws_state.write(0, col_idx, h, fmt_th)
    ws_state.set_column('A:D', 18)
    
    for row_idx, r in top_states.iterrows():
        ws_state.write(row_idx + 1, 0, r['customer_state'], fmt_td_text)
        ws_state.write(row_idx + 1, 1, r['revenue'], fmt_td_curr)
        ws_state.write(row_idx + 1, 2, r['orders'], fmt_td_int)
        ws_state.write(row_idx + 1, 3, r['avg_delivery'], fmt_td_dec)
        
    # -------------------------------------------------------------
    # SHEET 5: PAYMENTS & REVIEWS DATA
    # -------------------------------------------------------------
    ws_ops = workbook.add_worksheet('Data_Operations')
    ws_ops.write('A1', 'Payment Type', fmt_th)
    ws_ops.write('B1', 'Total Orders', fmt_th)
    ws_ops.write('C1', 'Total Revenue (R$)', fmt_th)
    for row_idx, r in payments.iterrows():
        ws_ops.write(row_idx + 1, 0, r['primary_payment_type'], fmt_td_text)
        ws_ops.write(row_idx + 1, 1, r['orders'], fmt_td_int)
        ws_ops.write(row_idx + 1, 2, r['revenue'], fmt_td_curr)
        
    ws_ops.write('E1', 'Review Rating', fmt_th)
    ws_ops.write('F1', 'Review Count', fmt_th)
    for row_idx, r in reviews.iterrows():
        ws_ops.write(row_idx + 1, 4, r['review_score_label'], fmt_td_text)
        ws_ops.write(row_idx + 1, 5, r['review_count'], fmt_td_int)
        
    ws_ops.write('H1', 'Delivery Timing vs Expectation', fmt_th)
    ws_ops.write('I1', 'Order Count', fmt_th)
    ws_ops.write('J1', 'Avg Review Score', fmt_th)
    for row_idx, r in delay_summary.iterrows():
        ws_ops.write(row_idx + 1, 7, str(r['delay_group']), fmt_td_text)
        ws_ops.write(row_idx + 1, 8, r['order_count'], fmt_td_int)
        ws_ops.write(row_idx + 1, 9, r['avg_review_score'], fmt_td_dec)
        
    ws_ops.set_column('A:C', 18)
    ws_ops.set_column('E:F', 18)
    ws_ops.set_column('H:J', 24)
    
    # -------------------------------------------------------------
    # CHARTS CREATION FOR EXECUTIVE DASHBOARD
    # -------------------------------------------------------------
    
    # 1. Line Chart: Monthly Revenue
    chart_revenue = workbook.add_chart({'type': 'line'})
    chart_revenue.add_series({
        'name': 'Monthly Revenue (R$)',
        'categories': ['Data_Monthly', 1, 0, len(monthly), 0],
        'values': ['Data_Monthly', 1, 1, len(monthly), 1],
        'line': {'color': '#0D9488', 'width': 2.75},
        'marker': {'type': 'circle', 'size': 5, 'color': '#0F172A'}
    })
    chart_revenue.set_title({'name': 'Monthly Revenue Growth (Jan 2017 – Aug 2018)', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_revenue.set_x_axis({'name_font': {'name': 'Segoe UI', 'size': 8}, 'num_font': {'name': 'Segoe UI', 'size': 8, 'rotation': -45}})
    chart_revenue.set_y_axis({'name': 'Revenue (R$)', 'major_gridlines': {'visible': True, 'line': {'color': '#F1F5F9'}}, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_revenue.set_legend({'position': 'none'})
    chart_revenue.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('B10', chart_revenue)
    
    # 2. Bar Chart: Top 10 Categories
    chart_cat = workbook.add_chart({'type': 'bar'})
    chart_cat.add_series({
        'name': 'Revenue (R$)',
        'categories': ['Data_Categories', 1, 0, len(top_cat), 0],
        'values': ['Data_Categories', 1, 1, len(top_cat), 1],
        'fill': {'color': '#1E293B'},
        'data_labels': {'value': True, 'num_font': {'name': 'Segoe UI', 'size': 7}}
    })
    chart_cat.set_title({'name': 'Top 10 Product Categories by Revenue', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_cat.set_y_axis({'reverse': True, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_cat.set_x_axis({'major_gridlines': {'visible': True, 'line': {'color': '#F1F5F9'}}, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_cat.set_legend({'position': 'none'})
    chart_cat.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('L10', chart_cat)
    
    # 3. Column Chart: Top States by Revenue
    chart_states = workbook.add_chart({'type': 'column'})
    chart_states.add_series({
        'name': 'State Revenue (R$)',
        'categories': ['Data_States', 1, 0, len(top_states), 0],
        'values': ['Data_States', 1, 1, len(top_states), 1],
        'fill': {'color': '#0D9488'}
    })
    chart_states.set_title({'name': 'Top 10 Brazilian States by Revenue', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_states.set_x_axis({'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_states.set_y_axis({'name': 'Revenue (R$)', 'major_gridlines': {'visible': True, 'line': {'color': '#F1F5F9'}}, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_states.set_legend({'position': 'none'})
    chart_states.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('B25', chart_states)
    
    # 4. Donut/Doughnut Chart: Payment Types
    chart_pay = workbook.add_chart({'type': 'doughnut'})
    chart_pay.add_series({
        'name': 'Orders by Payment Method',
        'categories': ['Data_Operations', 1, 0, len(payments), 0],
        'values': ['Data_Operations', 1, 1, len(payments), 1],
        'points': [
            {'fill': {'color': '#0F172A'}},
            {'fill': {'color': '#0D9488'}},
            {'fill': {'color': '#D97706'}},
            {'fill': {'color': '#94A3B8'}},
        ]
    })
    chart_pay.set_title({'name': 'Order Distribution by Payment Type', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_pay.set_legend({'position': 'right', 'font': {'name': 'Segoe UI', 'size': 8}})
    chart_pay.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('L25', chart_pay)
    
    # 5. Column Chart: Review Score Distribution
    chart_rev = workbook.add_chart({'type': 'column'})
    chart_rev.add_series({
        'name': 'Customer Reviews',
        'categories': ['Data_Operations', 1, 4, len(reviews), 4],
        'values': ['Data_Operations', 1, 5, len(reviews), 5],
        'fill': {'color': '#D97706'},
        'data_labels': {'value': True, 'num_font': {'name': 'Segoe UI', 'size': 7}}
    })
    chart_rev.set_title({'name': 'Customer Rating Breakdown (1–5 Stars)', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_rev.set_x_axis({'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_rev.set_y_axis({'major_gridlines': {'visible': True, 'line': {'color': '#F1F5F9'}}, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_rev.set_legend({'position': 'none'})
    chart_rev.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('B40', chart_rev)
    
    # 6. Column Chart: Delivery Delay vs Review Score
    chart_delay = workbook.add_chart({'type': 'column'})
    chart_delay.add_series({
        'name': 'Avg Rating',
        'categories': ['Data_Operations', 1, 7, len(delay_summary), 7],
        'values': ['Data_Operations', 1, 9, len(delay_summary), 9],
        'fill': {'color': '#0F172A'},
        'data_labels': {'value': True, 'num_format': '0.00', 'num_font': {'name': 'Segoe UI', 'size': 8, 'bold': True}}
    })
    chart_delay.set_title({'name': 'Impact of Delivery Delay on Customer Rating (1-5)', 'name_font': {'name': 'Segoe UI', 'size': 10, 'bold': True}})
    chart_delay.set_x_axis({'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_delay.set_y_axis({'min': 1, 'max': 5, 'major_gridlines': {'visible': True, 'line': {'color': '#F1F5F9'}}, 'num_font': {'name': 'Segoe UI', 'size': 8}})
    chart_delay.set_legend({'position': 'none'})
    chart_delay.set_size({'width': 500, 'height': 270})
    ws_dash.insert_chart('L40', chart_delay)
    
    workbook.close()
    print("Excel Dashboard successfully generated and saved!")

if __name__ == '__main__':
    create_excel_dashboard()
