# SKU Issue Analyzer 📊

A Python-based web application designed to automate the cleaning and analysis of e-commerce SKU issues.

This tool is built for Data Analysts and Customer Service teams to quickly generate reports from raw Order Excel files.

## 🚀 Key Features

* **Automated SKU Cleaning:** Detects and splits concatenated SKUs (e.g., `FR01FR02` → `FR01`, `FR02`).
* **Pivot Analysis:** Calculates total issues per SKU.
* **Smart Codification:** Assigns short codes (A, B, C...) to long issue descriptions for a cleaner summary.
* **Integrated Reporting:** Generates a single Excel sheet containing:
    1.  **Raw Cleaned Data**
    2.  **Summary Statistics** (Sorted by problem frequency)
    3.  **Detailed Legend** (With Total by SKU & Total by Unique Order metrics)

## 📂 How to Use

1.  Open the application (Deployed link).
2.  Upload your Order Excel file (`.xlsx`).
    * *Requirement:* The file must contain a sheet named **"Detail"**.
    * *Columns Used:* Column `C` (Order ID), `H` (Raw SKU), `J` (Reason).
3.  Click **Process Data**.
4.  Download the **Integrated_SKU_Report.xlsx**.

## 🛠️ Technology Stack

* **Python 3.x**
* **Streamlit** (Web Framework)
* **Pandas** (Data Manipulation)
* **XlsxWriter** (Excel Formatting)

## 💻 Local Installation

To run this app on your local machine:

```bash
# Clone the repository
git clone [https://github.com/your-username/sku-analyzer.git](https://github.com/your-username/sku-analyzer.git)

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
