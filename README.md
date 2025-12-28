
# JJM SWSM Daily Report Generator

A simple web app to generate the daily **JJM SWSM** report from the JJMUP export file.  
Upload the JJMUP export (`.xls` / `.xlsx`) and download an Excel report with two formatted sheets:
1) **SUPPLIED WATER LESS THAN 75**
2) **ZERO(INACTIVE SITES)**

---

## 🔗 Live App
✅ Streamlit App: https://jjmswsmdailyreport-bdubf94gkzgrz37plwhrww.streamlit.app/


---

## ✅ What this app produces

### Sheet 1: `SUPPLIED WATER LESS THAN 75`
Includes schemes where:

- **Percentage** = (Yesterday Water Production / Daily Water Demand) × 100  
- Percentage is **< 75%** (default threshold)

Columns:
- SR.No.
- Scheme Id
- Scheme Name
- Daily Water Demand (m^3)
- Yesterday Water Production (m^3)
- Percentage
- Supplied Water Percentage

> Note: “Today Water Production” is intentionally removed from this sheet.

---

### Sheet 2: `ZERO(INACTIVE SITES)`
Includes schemes where:
- **Yesterday Water Production = 0**
- **Today Water Production = 0**

Columns:
- SR.No.
- Scheme Id
- Scheme Name
- Yesterday Water Production (m^3)
- Today Water Production (m^3)
- Last Data Receive Date
- Site Status (ZERO/INACTIVE SITE)

---

## 🧾 How to use (for non‑technical users)

1. Open JJMUP website and export/download the daily report file (usually `.xls` or `.xlsx`).
2. Open the app link:
   https://jjmswsmdailyreport-bdubf94gkzgrz37plwhrww.streamlit.app/
3. Upload the exported file.
4. Click **Generate Report**.
5. Download the formatted Excel output.

---

## 📌 Important note about JJMUP `.xls`
Some JJMUP downloads are **HTML-based `.xls`** files (an HTML table saved with `.xls` extension).  
This app supports that format as well as normal `.xlsx`.

---

## 🖥️ Run locally (for developers)

### 1) Clone the repo
```bash
git clone https://github.com/Ashish-Vashisth/JJM_SWSM_DAILY_REPORT.git
cd JJM_SWSM_DAILY_REPORT
