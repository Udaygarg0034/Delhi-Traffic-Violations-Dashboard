# 🚦 Delhi Traffic Violations Dashboard — Excel Project

## 📌 Project Overview
An end-to-end **Excel Analytics Dashboard** built on Delhi Traffic Violations data containing **27,500+ violation records**. This project demonstrates real-world data modelling, KPI analysis, and interactive dashboard creation using Microsoft Excel.

---

## 📊 Dataset Structure — Star Schema

| Sheet | Type | Rows | Description |
|---|---|---|---|
| `fact_violation` | Fact Table | 27,500 | Core violations data — fine, speed, alcohol level, date |
| `dim_vehicle` | Dimension Table | 8,500 | Vehicle details — type, color, owner, license |
| `dim_officer` | Dimension Table | 135 | Officer details — rank, zone, agency |
| `KPI SHEET` | Analysis | — | Key Performance Indicators |
| `Pivot Table` | Analysis | — | Aggregated violation summaries |
| `DASHBOARD` | Visualization | — | Interactive Excel Dashboard |

---

## 🔑 Key KPIs Tracked

- Total Fine Amount Collected
- Total Violations Count
- Most Common Violation Type
- Top Violation Zones in Delhi
- Late Night Violations (9PM–3AM)
- Outside Delhi Vehicles Caught
- Speed Anomaly Detection
- Fine Payment Rate (%)

---

## 🛠️ Excel Skills & Formulas Used

### Lookup Functions
```excel
=VLOOKUP(B2, dim_vehicle!$A:$F, 6, 0)          -- Fetch Registration State
=VLOOKUP(C2, dim_officer!$A:$B, 2, 0)          -- Fetch Issuing Agency
```

### Conditional Logic
```excel
=IF(Recorded_Speed < Speed_Limit - 20, "Anomaly", "Normal")
=IF(Registration_State = "Delhi", "Delhi", "Outside Delhi")
```

### Date & Time Functions
```excel
=YEAR(Date)                                      -- Extract Year
=TEXT(Date, "MMM-YYYY")                          -- Format Month-Year
=IF(AND(HOUR(Time)>=21, HOUR(Time)<=23), "Late Night", "Normal")
```

---

## 📈 Dashboard Features

- ✅ Interactive Slicers for filtering by Zone, Agency, Vehicle Type
- ✅ Pivot Tables for violation type and fine amount analysis
- ✅ KPI Cards showing total fines, violations, payment rate
- ✅ Star Schema data model connecting 3 tables
- ✅ Speed Anomaly Detection using IF logic
- ✅ Time-based categorization (Late Night vs Normal)
- ✅ Month-Year trend analysis

---

## 🗂️ Project Structure

```
Delhi-Traffic-Violations-Dashboard/
│
├── Delhi_Traffic_Violations_DASHBOARD.xlsx    ← Main Excel File
└── README.md                                  ← Project Documentation
```

---

## 💡 Key Insights from Data

- **Over-speeding** is the most common violation in Delhi
- **Late Night hours (9PM–3AM)** show highest drunk driving cases
- **Outside Delhi vehicles** contribute significant fine revenue
- **Rohini, Dwarka, Lajpat Nagar** are top violation zones

---

## 🧰 Tools Used

| Tool | Purpose |
|---|---|
| Microsoft Excel | Data Modelling, Analysis, Dashboard |
| VLOOKUP | Joining dimension tables to fact table |
| Pivot Tables | Aggregation and summarization |
| IF / AND / OR | Conditional logic and flags |
| YEAR / TEXT / HOUR | Date-time feature engineering |

---

## 👤 Author

**Uday Garg**
- 📚 PG Program in Data Science & Analytics — Imarticus Learning
- 💼 Aspiring Data Analyst
- 🔗 [LinkedIn Profile](https://www.linkedin.com/in/uday-garg-b08374295)

---

## 📂 How to Use

1. Download the `.xlsx` file
2. Open in Microsoft Excel (2016 or later recommended)
3. Go to **DASHBOARD** sheet
4. Use slicers to filter data interactively
5. Check **KPI SHEET** for key metrics
