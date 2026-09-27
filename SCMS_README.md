# 📦 SCMS Supply Chain Analysis — Excel

## 📌 Project Overview

This project analyzes supply chain shipment data using Microsoft Excel to evaluate shipment performance, delivery reliability, freight costs, vendor performance, and country-level activity.

The workbook follows a practical analytics workflow:

**Raw Data → Data Cleaning → Analysis → Shipment Lookup → Interactive Dashboard**

The project was created as part of my analytics portfolio to demonstrate practical Excel and Supply Chain Analytics skills.

---

## 🎯 Business Objective

The objective is to understand supply chain performance and answer questions such as:

- How many shipments are being handled?
- What is the total shipment value?
- How long do shipments take to reach clients?
- What percentage of shipments are delivered on time?
- Which shipment modes are being used?
- How do vendors perform in terms of delivery and freight cost?
- Which countries have the highest shipment activity?
- Where are late deliveries occurring?

---

## 🗂️ Dataset

The workbook contains a supply chain shipment dataset with:

- **10,324 raw shipment records**
- **10,307 records after cleaning**
- Shipment, vendor, country, product, delivery, and freight-cost information

The workbook contains four main sheets:

1. `SCMS Dataset` — Original/raw data
2. `SCMS Dataset Cleaned` — Cleaned and prepared data
3. `SCMS Dataset Analysis` — Detailed analysis and lookup tools
4. `SCMS Dashboard` — Interactive supply chain performance dashboard

---

## 🧹 Data Cleaning

The raw dataset was reviewed and prepared for analysis.

The cleaned dataset contains **10,307 records**, compared with **10,324 records in the original dataset**.

The cleaned data was then used as the basis for the analysis and dashboard.

---

## 🛠️ Excel Skills & Techniques

This project demonstrates practical Excel techniques including:

- Data cleaning and preparation
- Excel Tables
- `XLOOKUP`
- `COUNTIF`
- `COUNTIFS`
- `SUMIF`
- `AVERAGEIF`
- `IFERROR`
- Calculated columns
- KPI calculations
- Pivot-style analysis
- Interactive filters
- Dashboard design
- Conditional analysis
- VBA-enabled workbook functionality

---

## 🔎 Shipment Lookup Tool

The workbook includes a shipment lookup tool that allows a user to enter a shipment ID and retrieve shipment-level information.

The lookup displays information such as:

- Shipment ID
- Country
- Vendor
- Product Group
- Shipment Mode
- Quantity
- Line Item Value
- Scheduled Delivery
- Delivered Date
- Delivery Status

This demonstrates how Excel can be used to create a practical operational lookup tool rather than only a reporting dashboard.

---

## 📊 Shipment Mode Analysis

The analysis compares shipment modes using:

- Shipment Count
- Total Shipment Value
- Average Delivery Days
- On-Time Delivery %
- Total Freight Cost
- Freight Cost %

This helps compare operational performance across different transportation modes.

---

## 🏢 Vendor Performance Analysis

Vendor-level analysis includes:

- Shipment Count
- Total Order Value
- Average Delivery Days
- Freight Cost
- On-Time %
- Average Freight Cost %
- Performance Rating

This provides a way to compare supplier/vendor performance using measurable operational indicators.

---

## 🌍 Country Performance Analysis

Country-level analysis includes:

- Shipment Count
- Total Value
- Average Delivery Days
- On-Time %
- Total Freight Cost
- Freight Cost %

This provides a geographic view of supply chain activity and delivery performance.

---

## 📈 Dashboard

The Supply Chain Performance Dashboard summarizes the overall operation using key KPIs and visualizations.

### Dashboard KPIs

The dashboard displays:

| KPI | Value |
|---|---:|
| Total Shipments | 10,307 |
| Total Shipment Value | $1.63B |
| Average Delivery Days | 106 Days |
| On-Time Delivery | 61.3% |
| Total Freight Cost | $68.65M |
| Countries Served | 43 |
| Total Vendors | 73 |
| Late Shipments | 1,185 |

The dashboard also includes:

- Shipment Mode Distribution
- Delivery Status Distribution
- Shipment Volume by Country
- Shipment Mode filters
- Vendor filters
- Country filters
- Delivery Status filters
- Product Group filters
- Reset Filters functionality

---

## 🖼️ Project Screenshots

### 📊 Supply Chain Performance Dashboard

![Supply Chain Dashboard](SCMS%20dashboard.png)

### 🔎 Shipment Lookup Tool

![Shipment Lookup Tool](SCMS%20shipment.png)

### 📈 Detailed Supply Chain Analysis

![Supply Chain Analysis](SCMS%20Analysis.png)

---

## 📁 Repository Structure

```text
scms-supply-chain-analysis/
│
├── README.md
├── SCMS Project.xlsm
├── SCMS Analysis.png
├── SCMS dashboard.png
└── SCMS shipment.png
```

---

## 💡 Key Business Insights

The dashboard provides a high-level view of the supply chain:

- The operation covers **10,307 shipments** across **43 countries** and **73 vendors**.
- Total shipment value is approximately **$1.63B**.
- Overall on-time delivery is **61.3%**, indicating that delivery reliability is an important performance area to monitor.
- There are **1,185 late shipments** in the analyzed dataset.
- Average delivery time is **106 days**.
- Total freight cost is approximately **$68.65M**.
- Shipment mode performance varies across shipment volume, delivery time, on-time rate, and freight cost.

These metrics can be used as starting points for deeper supply chain performance investigations.

---

## 🎯 Project Outcome

This project demonstrates how Excel can be used to transform raw supply chain data into:

**Clean Data → Operational Analysis → KPIs → Interactive Tools → Business Dashboard**

It is designed to demonstrate practical skills relevant to **Business Analyst, Supply Chain Analyst, and Data Analyst** roles.

---

## 📥 How to Use

1. Download `SCMS Project.xlsm`.
2. Open the workbook in Microsoft Excel.
3. Review the raw and cleaned datasets.
4. Explore the analysis sheet and shipment lookup tool.
5. Open the dashboard to interact with the filters and KPIs.

> **Note:** This is an `.xlsm` macro-enabled workbook, so Excel may require macros to be enabled for VBA-based functionality such as the reset/filter features.
