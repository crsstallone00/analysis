# Excel Data Cleaner (Streamlit)

A simple web app that cleans and classifies Excel data into **Clean** and **Dirty** sheets based on quality rules (missing key fields, invalid dates, phones, emails, age, salary, etc.).

Originally converted from a Google Colab script.

## Features
- Upload `.xlsx` / `.xls`
- Automatic quality checks (same logic as the original Colab version)
- Split into two sheets: `Clean_完整資料` and `Dirty_不完整資料`
- One-click download of the result

## Quick Start (Local)

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/excel-data-cleaner.git
cd excel-data-cleaner

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py