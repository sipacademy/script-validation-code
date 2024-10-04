import pandas as pd

# File paths
invoice_item_file = 'C:\\Users\\manmohan.d\\OneDrive - SAKSOFT LIMITED\\Desktop\\invoice_cal\\Invoice_item.xlsx'
invoice_file = 'C:\\Users\\manmohan.d\\OneDrive - SAKSOFT LIMITED\\Desktop\\invoice_cal\\Invoice.xlsx'

# Load both files
invoice_item_df = pd.read_excel(invoice_item_file, sheet_name='Invoice_Item')
invoice_df = pd.read_excel(invoice_file, sheet_name='Invoice')

# Define calculation functions
def calculate_group_totals(df):
    grouped_df = df.groupby(['Invoice Number', 'Center Code', 'Admission No']).agg(
        total_price=('price', 'sum'), # 
        total_amount=('amount', 'sum'), # basic amount
        total_tax=('tax', 'sum'), # total tax
        total_item=('total', 'sum') # invoice amount
    ).reset_index()
    return grouped_df

# Calculate group totals for invoice_item table
invoice_item_grouped = calculate_group_totals(invoice_item_df)

# Merge with the invoice table to compare aggregated values
merged_df = pd.merge(invoice_item_grouped, invoice_df, 
                     on=['Invoice Number', 'Center Code', 'Admission No'], 
                     suffixes=('_item', '_invoice'))

# Calculate total_invoice_amount
merged_df['total_invoice_amount'] = merged_df['total_amount'] - merged_df['DISCOUNT AMOUNT']

# Validation functions with updated logic and rounding
# validating invoice amount
def validate_invoice_amount(row):
    total_item_rounded = round(row['total_item'], 2)
    invoice_amount_rounded = round(row['INVOICE aMOUNT'], 2)
    return total_item_rounded == invoice_amount_rounded

#validating basic amount
def validate_basic_amount(row):
    total_amount_rounded = round(row['total_amount'], 2)
    total_basic_rounded = round(row['total_basic_amount'], 2)
    return total_amount_rounded == total_basic_rounded

# validating tax
def validate_tax(row):
    total_tax_rounded = round(row['total_tax'], 2)
    total_gst_rounded = round(row['TOTAL GST AMOUNT'], 2)
    return total_tax_rounded == total_gst_rounded

def validate_net_payable_amount(row):
    total_invoice_rounded = round(row['total_invoice_amount'], 2)
    calculated_net_payable = round(total_invoice_rounded + row['CGST (9%)'] + row['SGST (9%)'], 2)
    net_payable_rounded = round(row['net_payable_amount'], 2)
    return calculated_net_payable == net_payable_rounded

# Apply validation checks
merged_df['valid_invoice_amount'] = merged_df.apply(validate_invoice_amount, axis=1)
merged_df['valid_basic_amount'] = merged_df.apply(validate_basic_amount, axis=1)
merged_df['valid_tax_amount'] = merged_df.apply(validate_tax, axis=1)
merged_df['valid_net_payable'] = merged_df.apply(validate_net_payable_amount, axis=1)

# Load the invoice Excel file to add new columns
with pd.ExcelWriter(invoice_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    # Write the updated DataFrame to the 'Invoice' sheet
    merged_df.to_excel(writer, sheet_name='Invoice', index=False)

print("Validation results added and file saved.")
