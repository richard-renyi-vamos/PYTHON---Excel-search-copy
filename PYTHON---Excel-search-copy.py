import pandas as pd

# Load the Excel file
input_file = 'input.xlsx'  # Replace with your input file name
output_file = 'output.xlsx'  # Replace with your output file name

# Read the Excel file
df = pd.read_excel(input_file, sheet_name=None)

# Initialize an empty DataFrame to store the results
result_df = pd.DataFrame()

# Iterate over each sheet
for sheet_name, sheet_df in df.items():
    # Filter rows where any cell contains "CUSTOMER X"
    filtered_df = sheet_df[sheet_df.apply(lambda row: row.astype(str).str.contains('CUSTOMER X').any(), axis=1)]
    result_df = pd.concat([result_df, filtered_df])

# Write the filtered rows to a new Excel file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    result_df.to_excel(writer, index=False)

print(f'Filtered rows have been written to {output_file}')
