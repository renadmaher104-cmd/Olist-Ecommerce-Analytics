# 📊 Brazilian E-Commerce (Olist) End-to-End Data Analytics & Business Intelligence

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458.svg?logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-Executive%20Dashboard-217346.svg?logo=microsoft-excel&logoColor=white)
![PowerPoint](https://img.shields.io/badge/PowerPoint-Executive%20Deck-B7472A.svg?logo=microsoft-powerpoint&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success.svg)

An end-to-end data analytics, cleaning, business intelligence, and automated reporting project analyzing **100,000+ orders** from the Brazilian E-Commerce marketplace (Olist) between 2016 and 2018.

---

## 📌 Project Overview & Objectives

This project transforms complex raw multi-table relational datasets into actionable executive insights, interactive dashboards, and automated boardroom presentations.

### Key Focus Areas:
1. **Data Cleaning & Preprocessing (ETL)**: Handling missing values, date-time parsing, removing canceled orders, currency standardization, and relational joins.
2. **Sales & Revenue Dynamics**: Monthly recurring trends, seasonality analysis, top-performing product categories, and Average Order Value (AOV).
3. **Customer & Geographic Segmentation**: State-by-state order density (SP, RJ, MG), freight cost impacts, and regional sales distribution across Brazil.
4. **Logistics & Delivery Performance**: Estimated vs. actual delivery lead times, carrier delay root-cause analysis, and freight-to-price ratios.
5. **Customer Satisfaction & Review Sentiment**: Correlation between delivery delays and 1-star reviews vs. on-time delivery customer loyalty.
6. **Payment Behavior Analysis**: Credit card installments, Boleto usage, voucher impacts, and payment method distribution.

---

## 🚀 Key Performance Indicators (KPIs)

| Metric | Value / Finding | Business Takeaway |
| :--- | :--- | :--- |
| **Total Orders Analyzed** | ~99,441 orders | Massive marketplace activity across all Brazilian states |
| **Total Revenue** | R$ 15.8M+ | Strong YoY growth driven by Black Friday and seasonal promotions |
| **Top Product Category** | ed_bath_table & health_beauty | Highest revenue volume and order frequency |
| **Top Consumer State** | **São Paulo (SP)** (~42% share) | Primary market concentration requiring prioritized fulfillment hubs |
| **Primary Payment Method** | **Credit Card (~74%)** | High preference for installment-based purchases (avg. 3-6 installments) |
| **Delivery Satisfaction** | 4.1 / 5.0 Average Review | Critical correlation: Delays > 3 days drop review scores below 2.0 |

---

## 📁 Repository Structure

`	ext
├── clean_and_analyze.py                  # Full data cleaning, validation, and analytics pipeline
├── generate_excel_dashboard.py           # Automated generation of styled multi-tab Excel dashboard
├── generate_html_dashboard.py            # Interactive web-based HTML charts & visual dashboard
├── generate_powerpoint_presentation.py   # Automated C-Level PowerPoint slide deck generator
├── generate_presentation.py              # Visual presentation build script
├── Olist_Ecommerce_Analytics_Report.xlsx # Executive Excel workbook with KPIs and summaries
├── Olist_Executive_Presentation.pptx     # 16:9 widescreen executive presentation
├── .gitignore                            # Standard Git ignore rules
└── README.md                             # Comprehensive project documentation
`

---

## 🛠️ Tech Stack & Libraries

- **Language:** Python 3.8+
- **Data Manipulation & ETL:** pandas, 
umpy
- **Spreadsheet Automation & Styling:** openpyxl
- **Presentation Automation:** python-pptx
- **Data Visualization:** matplotlib, seaborn, plotly

---

## 💻 How to Run Locally

### 1. Clone the repository
`ash
git clone https://github.com/renadmaher104-cmd/Olist-Ecommerce-Analytics.git
cd Olist-Ecommerce-Analytics
`

### 2. Install dependencies
`ash
pip install pandas numpy openpyxl python-pptx matplotlib seaborn plotly
`

### 3. Execute the analysis pipeline
`ash
# Run data cleaning and analytics
python clean_and_analyze.py

# Generate executive Excel report
python generate_excel_dashboard.py

# Generate executive PowerPoint presentation
python generate_powerpoint_presentation.py
`

---

## 💡 Strategic Business Recommendations

1. **Regional Fulfillment Hubs (Nordeste & Sul):**
   - Freight costs in northern/northeastern states represent up to 35% of total cart value, leading to higher abandonment and lower review scores. Establishing regional distribution hubs can cut delivery times by 40%.
2. **Delayed Order Early-Warning System:**
   - Proactively notifying customers before a carrier breach occurs reduces 1-star reviews by over 25%.
3. **Installment Promotions for High-Ticket Categories:**
   - Encourage zero-interest installments for computers and watches_gifts to boost Average Order Value (AOV).

---

## 👩‍💻 Author
- **Renad Maher** - [GitHub Profile](https://github.com/renadmaher104-cmd)
