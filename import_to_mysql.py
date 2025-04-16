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

# 2. Read data from Excel
df = pd.read_excel("inventory_data.xlsx")

# 3. Loop through each row and insert
for _, row in df.iterrows():
    category = row['Category'].strip()
    brand = row['Brand'].strip()
    model = row['Model'].strip()
    quantity = int(row['Quantity'])
    log_date = datetime.today().date()

    # 3.1. Check if the device already exists
    cursor.execute("""
        SELECT id FROM devices
        WHERE category = %s AND brand = %s AND model = %s
    """, (category, brand, model))
    result = cursor.fetchone()

    if result:
        device_id = result[0]
    else:
        # If not found, insert a new device
        cursor.execute("""
            INSERT INTO devices (category, brand, model)
            VALUES (%s, %s, %s)
        """, (category, brand, model))
        conn.commit()
        device_id = cursor.lastrowid

    # 3.2. Insert inventory log for this device
    cursor.execute("""
        INSERT INTO inventory_logs (device_id, quantity, log_date)
        VALUES (%s, %s, %s)
    """, (device_id, quantity, log_date))

# 4. Commit and close connection
conn.commit()
cursor.close()
conn.close()

print("✅ Excel data imported to MySQL successfully!")
