
# IT Asset Inventory & Alert System

A complete inventory tracking and alert system for IT assets such as laptops, monitors, and accessories.

## Features
- Manage assets by category, brand, model, and quantity
- Automatically generate summary and low inventory reports (Excel)
- Email alert system when quantity falls below threshold
- Upload and track weekly burn rate of devices

## Technologies
- Python
- Pandas
- openpyxl (for Excel file handling)
- MySQL (for database backend)
- matplotlib (for chart visualization)
- smtplib (for sending emails)

##  How to Use

### 1. Import Inventory
```bash
python import_to_mysql.py
```
Loads inventory from `inventory_data.xlsx` into MySQL

### 2. Sync Current Inventory from Excel
```bash
python sync_excel_to_mysql.py
```
Syncs current quantities from `inventory_burn_rate.xlsm`

### 3. Upload Weekly Burn Rate
```bash
python import_weekly_burn_to_sql.py
```
Reads burn rate from `weekly_burn_rate.xlsx` and inserts to MySQL

### 4. Generate and Email Low Inventory Report
```bash
python send_alert.py
```
Generates `alert_report.xlsx` and sends email if low inventory is detected

##  Database Tables (MySQL)

### `devices`
| id | category   | brand | model |
|----|------------|-------|--------|

### `inventory_logs`
| id | device_id | quantity | log_date |

### `weekly_burn_rate`
| id | device_id | model | brand | category | status | week_start | week_end | burn_rate |

## 🛠 Installation

Install Python libraries:
```bash
pip install pandas openpyxl mysql-connector-python matplotlib
```

##  Email Setup
1. Enable 2-step verification in Gmail
2. Generate App Password
3. Replace it in `send_alert.py`

##  Future Plans
- Web interface using Streamlit
- REST API endpoints

## 👤 Author
Trieu VoVo 
System Analyst | Python Developer
