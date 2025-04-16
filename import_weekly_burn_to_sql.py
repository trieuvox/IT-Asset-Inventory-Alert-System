import pandas as pd
import mysql.connector
from datetime import datetime

# 1. Load Excel
df = pd.read_excel("weekly_burn_rate.xlsx")

# 2. Extract burn rate column (e.g., "04/14-04/18")
burn_column = [col for col in df.columns if '-' in col][0]
week_str = burn_column.strip()

# Convert to week_start and week_end
month1, day1 = map(int, week_str.split('-')[0].split('/'))
month2, day2 = map(int, week_str.split('-')[1].split('/'))
year = 2025  # or use datetime.today().year dynamically

week_start = datetime(year, month1, day1).date()
week_end = datetime(year, month2, day2).date()

# 3. Connect to MySQL
conn = mysql.connector.connect(
    host="127.0.0.1",
    user="root",
    password="Andrew#3105",
    database="inventory_system"
)
cursor = conn.cursor()

# 4. Insert each row
for _, row in df.iterrows():
    burn_rate = row[burn_column]
    if pd.isna(burn_rate):
        continue

    cursor.execute("""
        INSERT INTO weekly_burn_rate 
        (device_id, model, brand, category, status, week_start, week_end, burn_rate)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        row['ID'],
        row['Model'],
        row['Brand'],
        row['Category'],
        row['Status'],
        week_start,
        week_end,
        int(burn_rate)
    ))

# 5. Commit and close
conn.commit()
cursor.close()
conn.close()

print(f"✅ Uploaded weekly burn rate for {week_str} successfully!")
