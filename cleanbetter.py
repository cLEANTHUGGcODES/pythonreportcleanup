import pandas as pd
import datetime
import re
import matplotlib.pyplot as plt

input_file = r'C:\Users\James\Desktop\input.xlsx'

# Format the output file name with the current date
current_date = datetime.datetime.now().strftime('%m.%d.%Y')
output_file = fr'C:\Users\James\Desktop\OptumRx Accumulations {current_date}.xlsx'

# Read the Excel file
df = pd.read_excel(input_file)

# Drop columns by header name
columns_to_remove = ['Chain ID', 'Chain Name', 'HID', 'Hosp Name', 'Pharm Name', 'HRSA ID', 'Address', 'City', 'Pharm Effective', 'VID', 'Account Type']

for col in columns_to_remove:
    if col in df.columns:
        df.drop(col, axis=1, inplace=True)
        print(f"Removed column: {col}")
    else:
        print(f"Column not found: {col}")

# Remove columns B and C
df.drop(df.columns[[1, 2]], axis=1, inplace=True)

# Remove the last row
df.drop(df.tail(1).index, inplace=True)

# Sort the DataFrame so that rows with "Optum Frontier Therapies" come last within each group of duplicate PIDs
df = df.sort_values(by=['PID', 'Vendor Name'], ascending=[True, True])
df['Vendor Name'] = df['Vendor Name'].apply(lambda x: 'zzz' if x == 'Optum Frontier Therapies' else x)

# Remove rows with duplicate PID values, keeping the first occurrence
df = df.drop_duplicates(subset='PID', keep='first')

# Replace 'zzz' back to 'Optum Frontier Therapies' in the 'Vendor Name' column
df['Vendor Name'] = df['Vendor Name'].replace('zzz', 'Optum Frontier Therapies')

# Sort the DataFrame based on column A
df.sort_values(by=df.columns[0], inplace=True)

# Rename the column headers
df.rename(columns={'Utilization Count': '340B Accumulations', 'Total Accumulation Count': 'Total Utilization'}, inplace=True)

# Remove any non-numeric characters from the 'Account Number 340B' column
df['Account Number 340B'] = df['Account Number 340B'].astype(str).str.replace(r'\D', '', regex=True)

# Format the 'Terminated' column as a date field and replace the default date with "Not terminated"
default_date = '1900-01-01'
df['Terminated'] = pd.to_datetime(df['Terminated'], errors='coerce').apply(lambda x: 'Not terminated' if x is pd.NaT or x.strftime('%Y-%m-%d') == default_date else x)

# Save the modified DataFrame to a new Excel file with custom formatting
writer = pd.ExcelWriter(output_file, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}})
df.to_excel(writer, index=False, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Define the formats
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#4472C4',
    'border': 1,
    'font_color': 'white'})

cell_format = workbook.add_format({
    'border': 1,
    'align': 'left',
    'valign': 'top',
    'text_wrap': True})

# Apply the formats
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0, col_num, value, header_format)

for row_num, row_data in enumerate(df.values):
    for col_num, col_data in enumerate(row_data):
        worksheet.write(row_num + 1, col_num, col_data, cell_format)

# Autofit the columns
for idx, col in enumerate(df):
    series = df[col]
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
    )) + 1  # adding a little extra space
    worksheet.set_column(idx, idx, max_len)

# Calculate vendor_counts after removing duplicates
vendor_counts = df['Vendor Name'].value_counts()

# Remove specified vendors from vendor_counts
vendors_to_exclude = ['Optum Frontier Therapies', 'MCKESSON DROPSHIP']
vendor_counts = vendor_counts[~vendor_counts.index.isin(vendors_to_exclude)]

# Create a pie chart of the vendor breakdown with a legend
fig, ax = plt.subplots()
wedges, texts, autotexts = ax.pie(vendor_counts, autopct='%1.1f%%', startangle=90)
ax.axis('equal')  # Equal aspect ratio ensures the pie chart is circular
plt.title('Vendor Breakdown by PID (Excluding Optum Frontier Therapies)')

# Add a legend to the chart
ax.legend(wedges, vendor_counts.index, title='Vendors', loc='center left', bbox_to_anchor=(1, 0, 0.5, 1))

# Save the pie chart as an image file
plt.savefig(fr'C:\Users\James\Desktop\Vendor_Breakdown_No_Optum_{current_date}.png', dpi=300, bbox_inches='tight')

# Close the writer and save the file
writer.close()

