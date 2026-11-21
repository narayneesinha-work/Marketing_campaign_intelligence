
# Marketing Campaign Intelligence System

**An interactive marketing analytics application built with Python and Streamlit to evaluate campaign performance, analyse conversion funnels, and identify potential performance issues.**

## 🚀 Live Application

[Open Marketing Campaign Intelligence](https://marketingcampaignintelligence-8hjerqdzwssdskyacgfkhsq.streamlit.app/)

## 📌 Project Overview

Marketing teams generate large volumes of advertising data across campaigns, ad sets, advertisements, and audience segments. Analysing these reports separately can make it difficult to understand overall performance and identify where improvements are needed.

The Marketing Campaign Intelligence System simplifies this process by consolidating available campaign reports into an interactive dashboard and generating evidence-based performance insights.

## ✨ Key Features

### 1. Campaign Performance Dashboard
- Monitor total spend, impressions, clicks, purchases, and revenue.
- Analyse CTR, CPC, CPM, CPA, and ROAS.
- Rank campaigns from highest to lowest performance.
- Identify high-spend campaigns with relatively weaker returns.

### 2. Conversion Funnel Analysis
- Analyse the journey from clicks to landing page views, add to carts, checkouts, and purchases.
- Identify conversion drop-offs between funnel stages.
- Compare funnel performance across campaigns.

### 3. Campaign-Level Diagnostics
- Identify stronger and weaker campaign performers.
- Compare performance against selected-data benchmarks.
- Highlight metrics that may require investigation.
- Generate potential explanations supported by available data.

### 4. Master Data Analysis
- Consolidate available campaign reports.
- Filter and explore campaign, ad set, and ad-level data.
- Export analysed data for further examination.

### 5. Data Quality & Audit
- Validate report structures and date fields.
- Check duplicate records and merge keys.
- Identify missing values and unmatched records.
- Preserve unavailable metrics rather than treating them as zero.

## 🛠️ Technology Stack

- Python
- Streamlit
- Pandas
- Plotly
- OpenPyXL
- GitHub

## 📊 How to Use

1. Open the live application.
2. Select the Meta Ads platform.
3. Upload your campaign Excel workbook.
4. Explore the performance dashboard.
5. Apply date and campaign filters.
6. Review funnel analysis and campaign diagnostics.
7. Explore master data and download analysis reports.

## 🔍 Project Objective

The objective is to move beyond basic advertising dashboards and build a practical campaign intelligence tool that helps marketers interpret performance data, investigate conversion bottlenecks, and make more informed optimisation decisions.

## ⚠️ Important Notes

- The current implementation supports Meta Ads workbook analysis.
- Analysis depends on the metrics and reports supplied in the uploaded workbook.
- Missing metrics are treated as unavailable, not as zero.
- Performance flags are analytical indicators, not proof of causation or profitability.

## 👩‍💻 Project Status

**Live deployment available.**

This project demonstrates practical applications of Python-based marketing analytics, data processing, campaign performance measurement, and interactive reporting.
