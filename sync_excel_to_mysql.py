import pandas as pd
import mysql.connector
from datetime import datetime

# 1. Connect to MySQL
conn = mysql.connector.connect(
    host="127.0.0.1",
    user="root",
    password="Andrew#3105",
    database="inventory_system"
)
cursor = conn.cursor()

# 2. Read Excel file
df = pd.read_excel("inventory_burn_rate.xlsm", sheet_name="Data summary",engine="openpyxl")

# 3. Loop through each row of data
today = datetime.today().date()

for _, row in df.iterrows():
    brand = str(row['Brand']).strip()
    model = str(row['Model']).strip()
    category = str(row['Category']).strip()
    quantity = int(row['Current Quantity'])

    # 3.1. Check if the device already exists
    cursor.execute("""
        SELECT id FROM devices
        WHERE category = %s AND brand = %s AND model = %s
    """, (category, brand, model))
    result = cursor.fetchone()

    if result:
        device_id = result[0]
    else:
        # Insert the device if it doesn't exist
        cursor.execute("""
            INSERT INTO devices (category, brand, model)
            VALUES (%s, %s, %s)
        """, (category, brand, model))
        conn.commit()
        device_id = cursor.lastrowid

    # 3.2. Insert the current inventory log
    cursor.execute("""
        INSERT INTO inventory_logs (device_id, quantity, log_date)
        VALUES (%s, %s, %s)
    """, (device_id, quantity, today))

# 4. Commit and close connection
conn.commit()
cursor.close()
conn.close()

print("✅ Synced Excel to MySQL successfully!")
