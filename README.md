# 📊 Brazilian E-Commerce (Olist) Data Analytics & Executive Reporting

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458.svg?logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-Analytics%20Report-217346.svg?logo=microsoft-excel&logoColor=white)
![PowerPoint](https://img.shields.io/badge/PowerPoint-Executive%20Presentation-B7472A.svg?logo=microsoft-powerpoint&logoColor=white)
![HTML5](https://img.shields.io/badge/Interactive-HTML%20Deck-E34F26.svg?logo=html5&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success.svg)

An end-to-end data analytics, cleaning, business intelligence, and executive reporting project analyzing **100,000+ customer orders** from the Brazilian E-Commerce marketplace (Olist).

---

## 📌 Project Overview

This project provides comprehensive business insights, performance analytics, and data-driven recommendations across core commercial functions:
- **Sales & Revenue Performance**: Revenue growth trajectories, seasonality peaks (e.g., Black Friday), Average Order Value (AOV), and top-grossing product categories.
- **Customer & Geographic Dynamics**: Order density across Brazilian states (São Paulo, Rio de Janeiro, Minas Gerais), freight costs, and regional demand.
- **Logistics & Delivery Efficiency**: Delivery lead times, freight-to-price ratios, carrier delay impacts, and estimated vs. actual arrival times.
- **Review Scores & Customer Satisfaction**: Analyzing the root causes of negative reviews (1-star) and their strong correlation with shipping delays.
- **Payment Trends**: Distribution of payment methods (Credit Card, Boleto, Voucher, Debit Card) and installment behavior.

---

## 🚀 Key Performance Indicators (KPIs)

| Metric | Insight / Value | Strategic Business Takeaway |
| :--- | :--- | :--- |
| **Total Orders Analyzed** | **~99,441 orders** | Broad market coverage across 27 Brazilian states |
| **Total Revenue** | **R$ 15.8M+** | Substantial revenue driven by top retail categories |
| **Top Category by Sales** | `bed_bath_table`, `health_beauty`, `watches_gifts` | High volume and high-margin product categories |
| **Geographic Core** | **São Paulo (SP)** (~42% of all orders) | Concentrated hub requiring optimized fulfillment centers |
| **Payment Preference** | **Credit Card (~74%)** | Customers heavily rely on installment payment plans (3–6 months) |
| **Customer Satisfaction** | **4.1 / 5.0 Average Review** | Delivery delays > 3 days cause 1-star ratings to surge |

---

## 📁 Repository Contents

```text
├── clean_and_analyze.py                  # End-to-end data cleaning, preprocessing & analytics script
├── Olist_Ecommerce_Analytics_Report.xlsx # Comprehensive executive Excel report with KPI summaries
├── Olist_Executive_Presentation.pptx     # 16:9 Widescreen C-Level PowerPoint presentation
├── presentation.html                     # Interactive bilingual HTML presentation dashboard
├── .gitignore                            # Standard Git ignore configuration
└── README.md                             # Full project documentation & insights
```

### 📄 Detailed File Descriptions:
1. **`clean_and_analyze.py`**:
   - Performs complete ETL pipeline: cleans raw datasets, handles missing values, converts timestamps, filters canceled transactions, and computes key business metrics.
2. **`Olist_Ecommerce_Analytics_Report.xlsx`**:
   - Multi-sheet executive workbook presenting sales breakdown, category metrics, state-by-state analysis, payment insights, and delivery KPIs.
3. **`Olist_Executive_Presentation.pptx`**:
   - Professionally formatted executive presentation ready for stakeholder review and boardroom discussions.
4. **`presentation.html`**:
   - Modern, interactive HTML presentation deck designed for seamless browser demonstration.

---

## 🛠️ Tools & Technologies

- **Programming Language:** Python
- **Data Analysis & Processing:** Pandas, NumPy
- **Reporting & Dashboards:** Microsoft Excel, OpenPyXL, HTML5/CSS3/JS
- **Presentation Deck:** Microsoft PowerPoint (python-pptx)

---

## 💡 Key Strategic Recommendations

1. **Optimize Regional Fulfillment (Nordeste & Sul):**
   - Freight costs in remote states can reach up to 35% of total order value. Establishing regional micro-fulfillment centers will reduce transit times by up to 40% and lower cancellation rates.
2. **Proactive Delay Alert System:**
   - Automatically notify buyers whenever a carrier delay is detected to mitigate negative reviews and maintain high satisfaction scores.
3. **Targeted Installment Campaigns:**
   - Offer zero-interest installment promotions on high-ticket items (`computers`, `watches_gifts`) to maximize Average Order Value (AOV).

---

## 👩‍💻 Author
- **Renad Maher** - [GitHub Profile](https://github.com/renadmaher104-cmd)
