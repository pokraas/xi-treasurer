# AI-generated bad code


import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Load the CSV file
df = pd.read_csv('data/stmt october.csv')

# Load the Excel workbook
book = load_workbook('data/Xi 2023-24 02_September.xlsx')

# Get the current sheet
current_sheet = book['Current']

# Check the beginning balance
if df['Beginning balance'].iloc[0] != current_sheet['G5'].value:
    print("Beginning balance does not match!")
    exit(1)

# Update the date, month, beginning balance date and amount
current_sheet['A1'] = df['Date'].iloc[0]
current_sheet['B1'] = df['Month'].iloc[0]
current_sheet['C1'] = df['Beginning balance Date'].iloc[0]
current_sheet['D1'] = df['Amount'].iloc[0]

# Switch to the Credits tab
credits_sheet = book['Credits']

# Sort the transactions by amount
df = df.sort_values(by='Amount', ascending=False)

# Copy the credit rows to the Credits tab
for index, row in df.iterrows():
    credits_sheet.append(row.values)

# Save the workbook
book.save('output.xlsx')
