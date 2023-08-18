import pandas as pd

# Replace 'your_file.xlsx' with the actual path to your Excel file
excel_file = 'Input.xlsx'

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file)

# Replace 'voucher type' with the actual column name containing voucher types
voucher_type_column = 'VOUCHERTYPENAME'
target_voucher_type = 'Receipt'

# Filter the DataFrame to include only rows with the target voucher type
filtered_df = df[df[voucher_type_column] == target_voucher_type]

# Replace ['column1', 'column2'] with the actual column names you want to extract
columns_to_extract = ['DATE', 'REFERENCEDATE', 'NARRATION','VOUCHERTYPENAME','REFERENCE','VOUCHERNUMBER','PARTYLEDGERNAME']


# Create a DataFrame with only the selected columns
selected_columns_df = filtered_df[columns_to_extract]

# Replace 'output_file.xlsx' with the desired name for the output Excel file
output_excel_file = 'output_file4.xlsx'

# Save the selected columns DataFrame to a new Excel file
selected_columns_df.to_excel(output_excel_file, index=False)

print(f"Data saved to {output_excel_file}")
