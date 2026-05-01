# Project 2: Excel Consolidation + Dashboard
# Day 11 - Jessie Portfolio
# Combine 3 monthly sales files into 1 master report with summary
import pandas as pd
import glob

jan = pd.DataFrame({
    'Product': ['iPhone15', 'Laptop'],
    'Price': [70000, 35000],
    'Month': ['Jan', 'Jan']
})
jan.to_excel('Sales_Jan.xlsx', index=False)
feb = pd.DataFrame({
    'Product': ['iPad', 'MacBook'],
    'Price': [25000, 80000],
    'Month': ['Feb', 'Feb']
})
feb.to_excel('Sales_Feb.xlsx', index=False)
mar = pd.DataFrame({
    'Product': ['Airpods', 'iPhone15'],
    'Price': [10000, 70000],
    'Month': ['Mar', 'Mar']
})
mar.to_excel('Sales_Mar.xlsx', index=False)
print("✅ 3 sample files ginawa: Sales_Jan.xlsx, Sales_Feb.xlsx, Sales_Mar.xlsx")
# Start processing
excel_files = glob.glob('Sales_*.xlsx')
print("NA-DETECT NA FILES:", excel_files)
all_data = []
for file in excel_files:
    df = pd.read_excel(file)
    all_data.append(df)
master_file = pd.concat(all_data, ignore_index=True)
print("MASTER FILE - LAHAT NG DATA:")
print(master_file)
summary = master_file.groupby('Month')['Price'].sum().reset_index()
summary.columns = ['Month', 'Total Sales']
print("TOTAL PER MONTH:")
print(summary)
filename = 'MASTER_SALES_REPORT.xlsx'
with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    master_file.to_excel(writer, sheet_name='All Data', index=False)
    summary.to_excel(writer, sheet_name='Monthly Total', index=False)
    for sheet in writer.sheets:
        worksheet = writer.sheets[sheet]
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 18
        for row in range(2, worksheet.max_row + 1):
            worksheet[f'B{row}'].number_format = '₱#,##0'
print(f"\n✅ MASTER FILE WITH 2 SHEETS SAVED AS {filename}")
