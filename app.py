# ============================================================
# 完整 Colab 程式碼：上傳 → 清理 → 分 Sheet → 自動下載
# ============================================================

from google.colab import files
import pandas as pd
import numpy as np
import re
from datetime import datetime
import io

print("請點擊下方按鈕上傳 Excel 檔案...")
uploaded = files.upload()

# 讀取上傳的檔案（取第一個檔案）
filename = list(uploaded.keys())[0]
df = pd.read_excel(io.BytesIO(uploaded[filename]))

print(f"\n成功讀取檔案：{filename}")
print(f"原始資料筆數：{len(df)}")
print(f"欄位數量：{len(df.columns)}")
print("-" * 50)

# ============================================================
# 資料清理與品質判斷函式
# ============================================================

def is_null_or_empty(value):
    """判斷是否為空值、NULL、N/A、None、nan 等"""
    if pd.isna(value):
        return True
    if isinstance(value, str):
        val = value.strip().upper()
        if val in ['', 'NULL', 'N/A', 'NA', 'NONE', '#NA', 'NIL', 'UNKNOWN', '-', '#N/A']:
            return True
    return False

def is_valid_date(value):
    """檢查日期格式是否合理"""
    if is_null_or_empty(value):
        return False
    try:
        # 嘗試多種常見格式
        if isinstance(value, (pd.Timestamp, datetime)):
            year = value.year
            return 1950 <= year <= 2030
        val_str = str(value).strip()
        # 拒絕明顯錯誤的日期
        if re.search(r'1990-00-00|2023-02-31|0000|INVALID', val_str, re.IGNORECASE):
            return False
        # 嘗試解析
        parsed = pd.to_datetime(val_str, errors='coerce', dayfirst=True)
        if pd.isna(parsed):
            return False
        return 1950 <= parsed.year <= 2030
    except:
        return False

def is_valid_phone(value):
    """檢查電話號碼格式是否合理（針對澳門 +853 為主）"""
    if is_null_or_empty(value):
        return False
    val = str(value).strip()
    # 移除常見符號後檢查
    cleaned = re.sub(r'[\s\-\(\)\+]', '', val)
    # 至少要有 8 位數字，且不能全是 0 或明顯無效
    if not re.search(r'\d{8,}', cleaned):
        return False
    if re.search(r'12345|00000|99999|NULL|N/A', val, re.IGNORECASE):
        return False
    return True

def is_valid_email(value):
    """簡單檢查 Email 格式"""
    if is_null_or_empty(value):
        return False
    val = str(value).strip()
    # 拒絕明顯錯誤的 email
    if re.search(r'@@|\.\.|none_provided|invalid_user|DOMAIN\.ORG', val, re.IGNORECASE):
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, val))

def is_valid_age(value):
    """檢查年齡是否合理"""
    if is_null_or_empty(value):
        return False
    try:
        age = float(str(value).replace(',', ''))
        return 16 <= age <= 80
    except:
        return False

def is_valid_salary(value):
    """檢查薪資是否合理"""
    if is_null_or_empty(value):
        return False
    try:
        salary = float(str(value).replace(',', '').replace('$', '').replace('HK$', '').replace('MOP', '').strip())
        return 5000 <= salary <= 100000   # 合理範圍
    except:
        return False

def has_mixed_case_or_messy_name(value):
    """檢查姓名是否有明顯大小寫混亂或空值"""
    if is_null_or_empty(value):
        return True
    val = str(value).strip()
    # 全大寫或全小寫且長度較長視為可能問題，或有多個空格
    if re.search(r'\s{2,}', val):
        return True
    return False

def is_dirty_row(row):
    """綜合判斷一筆資料是否為 Dirty"""
    # 1. 關鍵欄位空值檢查
    key_fields = ['First_Name', 'Last_Name', 'Email_Address', 'Phone_Number',
                  'Department', 'Job_Title', 'Employment_Status', 'Salary_Monthly']
    for field in key_fields:
        if field in row and is_null_or_empty(row[field]):
            return True

    # 2. 日期格式檢查
    date_fields = ['DOB', 'Hire_Date', 'Contract_End_Date', 'Registration_Date',
                   'Last_Login_Timestamp', 'Last_Purchase_Date']
    for field in date_fields:
        if field in row and not is_null_or_empty(row[field]):
            if not is_valid_date(row[field]):
                return True

    # 3. 電話格式
    if 'Phone_Number' in row and not is_valid_phone(row['Phone_Number']):
        return True
    if 'Emergency_Contact_Phone' in row and not is_null_or_empty(row['Emergency_Contact_Phone']):
        if not is_valid_phone(row['Emergency_Contact_Phone']):
            return True

    # 4. Email 格式
    if 'Email_Address' in row and not is_valid_email(row['Email_Address']):
        return True

    # 5. 年齡異常
    if 'Age' in row and not is_valid_age(row['Age']):
        return True

    # 6. 薪資異常
    if 'Salary_Monthly' in row and not is_valid_salary(row['Salary_Monthly']):
        return True

    # 7. 姓名明顯混亂
    if 'First_Name' in row and has_mixed_case_or_messy_name(row['First_Name']):
        return True
    if 'Last_Name' in row and has_mixed_case_or_messy_name(row['Last_Name']):
        return True

    # 8. 如果原本就有 Data_Quality_Label 欄位，直接參考
    if 'Data_Quality_Label' in row:
        label = str(row['Data_Quality_Label']).strip().upper()
        if label == 'DIRTY':
            return True

    # 9. IP / MAC 明顯無效
    if 'IP_Address' in row:
        ip = str(row['IP_Address']).strip()
        if re.search(r'999\.999\.999\.999|localhost|10\.0\.0\.1\.5|N/A|NULL', ip, re.IGNORECASE):
            return True

    if 'MAC_Address' in row:
        mac = str(row['MAC_Address']).strip()
        if re.search(r'INVALID_MAC|NULL|N/A', mac, re.IGNORECASE):
            return True

    return False

# ============================================================
# 執行分類
# ============================================================

print("開始判斷資料品質...")

dirty_mask = df.apply(is_dirty_row, axis=1)

df_clean = df[~dirty_mask].copy()
df_dirty = df[dirty_mask].copy()

print(f"完整資料 (Clean)：{len(df_clean)} 筆")
print(f"不完整資料 (Dirty)：{len(df_dirty)} 筆")
print("-" * 50)

# ============================================================
# 輸出整理後的 Excel
# ============================================================

output_filename = "整理後_Clean_Dirty_資料.xlsx"

with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    df_clean.to_excel(writer, sheet_name='Clean_完整資料', index=False)
    df_dirty.to_excel(writer, sheet_name='Dirty_不完整資料', index=False)

print(f"已產生檔案：{output_filename}")
print("正在自動下載到您的本機電腦...")

files.download(output_filename)

print("\n✅ 全部完成！檔案已下載。")
print("檔案內含兩個工作表：")
print("  1. Clean_完整資料")
print("  2. Dirty_不完整資料")

convert the above code into github and steamlit