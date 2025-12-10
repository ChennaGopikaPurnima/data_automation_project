import pandas as pd
import glob
import os

# PATHS
DATA_PATH = r"C:\Users\chenn\Desktop\data_automation_project\data"
OUTPUT_PATH = r"C:\Users\chenn\Desktop\data_automation_project\output\Final_Report.xlsx"

# 1. Read all sales region files
files = glob.glob(os.path.join(DATA_PATH, "sales_region*.xlsx"))
print("Files found:", files)

if not files:
    print("❌ No sales_region files found!")
    exit()

# 2. Combine all files
dfs = [pd.read_excel(f) for f in files]
sales = pd.concat(dfs, ignore_index=True)

# 3. Read product master
master = pd.read_excel(os.path.join(DATA_PATH, "product_master.xlsx"))

# 4. Merge sales + master
final = pd.merge(sales, master, on="Product", how="left")

# 5. Add computed column
final["Total_Sales"] = final["Qty"] * final["Price"]

# 6. Remove duplicates
final.drop_duplicates(inplace=True)

# 7. Save final report
final.to_excel(OUTPUT_PATH, index=False)

print("✅ Automation Completed! Report created:")
print(OUTPUT_PATH)
