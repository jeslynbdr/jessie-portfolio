# Project 1: Data Cleaning Service
# Day 10 - Jessie Portfolio
# Clean messy sales data: remove ₱, commas, blanks, duplicates

import pandas as pd

# Sample dirty data
dirty_data = pd.DataFrame({
    'Product': ['iPhone15', 'Laptop', 'iPad', 'MacBook', 'Airpods', 'iPhone15'],
    'Price': ['₱70,000', '35000', '25,000.00', None, '₱10,000', '₱70,000']
})

print("BEFORE - CLIENT'S MESSY DATA:")
print(dirty_data)

# Start cleaning
clean = dirty_data.copy()

# 1. Remove ₱ and comma from Price column
clean['Price'] = clean['Price'].astype(str).str.replace('₱','').str.replace(',','')

# 2. Convert to numbers, fill blanks with 0
clean['Price'] = pd.to_numeric(clean['Price'], errors='coerce')
clean['Price'] = clean['Price'].fillna(0).astype(int)

# 3. Remove duplicate products, keep first entry
clean = clean.drop_duplicates(subset=['Product'], keep='first')

# 4. Add TOTAL row
total = clean['Price'].sum()
clean.loc[len(clean)] = ['TOTAL', total]

# 5. Save to Excel with Peso format
filename = 'CLEAN_REPORT.xlsx'
with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    clean.to_excel(writer, sheet_name='Clean Data', index=False)
    worksheet = writer.sheets['Clean Data']
    for row in range(2, worksheet.max_row + 1):
        worksheet[f'B{row}'].number_format = '₱#,##0'

print("\nAFTER - CLEAN DATA:")
print(clean)
print(f"\n✅ Saved to {filename}")
