from datetime import datetime
import pandas as pd
from thefuzz import fuzz, process
import glob
import os
import re


folder_path = (r"C:\xxxxxxxxxxxxxxxxxxxxxx\duplicates")
keyword = "Invoice Details for Duplicates"

# Searching for all Excel files that contain the keyword in their name
file_paths = glob.glob(os.path.join(folder_path, f"*{keyword}*.xlsx"))


dataframes = []

# Looping through a list of file paths, attempts to load each CSV file into a pandas DataFrame, 
# stores the DataFrames in a list, and handles errors if a file cannot be loaded.

for file_path in file_paths:
    try:
        df = pd.read_excel(file_path, header =0)
        dataframes.append(df)
        print(f'Loaded file: {file_path}')
    except Exception as e:
        print(f'Error loading file {file_path}: {e}')

# Combining all DataFrames into one
combined_df = pd.concat(dataframes, ignore_index=True)

# Filtering out rows where 'Invoice Amount' is not equal to 0.0 (it means that the invoice was cancelled in the ERP.)
combined_df = combined_df[combined_df['Invoice Amount'] != 0.0]


# FUNCTIONS

# Prefiltering the data for Incorrect Supplier Check
def pre_filter_supplier(combined_df):
    return combined_df[combined_df.duplicated(subset=['Invoice Date', 'Legal Entity', 'Invoice Amount', 'Business Unit'], keep=False)]

# Prefiltering the data for Incorrect Legal Entity Check
def pre_filter_le(combined_df):
    return combined_df[combined_df.duplicated(subset=['Supplier Name', 'Invoice Date', 'Invoice Amount', 'Business Unit'], keep=False)]

# Normalizing the invoice number: converting to uppercase and removing special characters
def normalize_invoice_number(number):
    return re.sub(r'\W+', '', str(number)).upper()

# Function that compares 2 invoice numbers. Threshold may be adjusted (95 = 95% of similarity)
def is_similar(s1, s2, threshold=95):
    s1_normalized = normalize_invoice_number(s1)
    s2_normalized = normalize_invoice_number(s2)
    return fuzz.ratio(s1_normalized, s2_normalized) >= threshold


# Detects potential duplicate invoices across different suppliers by comparing similar invoice numbers within groups 
# defined by invoice date, legal entity, invoice amount, and business unit. Ignoring Supplier name - This helps identify 
# potential duplicate or suspicious invoices that share the same key attributes but were posted as from different vendors.

def find_potential_duplicates_supplier(combined_df):
    combined_df['Invoice Number'] = combined_df['Invoice Number'].astype(str)
    grouped = combined_df.groupby(['Invoice Date', 'Legal Entity', 'Invoice Amount', 'Business Unit'])
    
    duplicates = []
    
    for _, group in grouped:
        invoice_numbers = group['Invoice Number'].tolist()
        business_unit_names = group['Business Unit'].tolist()
        supplier_names = group['Supplier Name'].tolist()
        inv_amounts = group['Invoice Amount'].tolist()

        for i in range(len(invoice_numbers)):
            for j in range(i + 1, len(invoice_numbers)):
                if supplier_names[i] != supplier_names[j] and is_similar(invoice_numbers[i], invoice_numbers[j]):
                    
                    duplicates.append({
                        'Invoice Number 1': invoice_numbers[i],
                        'Invoice Number 2': invoice_numbers[j],
                        'Supplier 1': supplier_names[i],
                        'Supplier 2': supplier_names[j],
                        'Amount': inv_amounts[i],
                        'Business Unit': business_unit_names[i]
                    })
    
    return pd.DataFrame(duplicates)

# Detects potential duplicate invoices across different suppliers by comparing similar invoice numbers within groups 
# defined by invoice date, supplier name, invoice amount, and business unit. Ignoring Legal Entity - This helps identify 
# potential duplicate or suspicious invoices that share the same key attributes but were posted on different legal entities.

def find_potential_duplicates_le(combined_df):
    combined_df['Invoice Number'] = combined_df['Invoice Number'].astype(str)
    grouped = combined_df.groupby(['Supplier Name', 'Invoice Date', 'Invoice Amount', 'Business Unit'])
    
    duplicates = []
    
    for _, group in grouped:
        invoice_numbers = group['Invoice Number'].tolist()
        business_unit_names = group['Business Unit'].tolist()
        supplier_names = group['Supplier Name'].tolist()
        inv_amounts = group['Invoice Amount'].tolist()

        for i in range(len(invoice_numbers)):
            for j in range(i + 1, len(invoice_numbers)):
                if is_similar(invoice_numbers[i], invoice_numbers[j]):
                    
                    duplicates.append({
                        'Invoice Number 1': invoice_numbers[i],
                        'Invoice Number 2': invoice_numbers[j],
                        'Supplier 1': supplier_names[i],
                        'Supplier 2': supplier_names[j],
                        'Amount': inv_amounts[i],
                        'Business Unit': business_unit_names[i]
                    })
    
    return pd.DataFrame(duplicates)

filtered_df_supplier = pre_filter_supplier(combined_df)
potential_duplicates_supplier = find_potential_duplicates_supplier(filtered_df_supplier)

filtered_df_le = pre_filter_le(combined_df)
potential_duplicates_le = find_potential_duplicates_le(filtered_df_le)

# Function that saves the results to a history file to avoid detecting the same duplicates in future runs
def save_new_and_history(potential_duplicates, sheet_name, new_file_path, history_file_path):
    today = datetime.today().strftime('%Y-%m-%d')
    potential_duplicates['Found Date'] = today

    if os.path.exists(history_file_path):
        try:
            history_df = pd.read_excel(history_file_path, sheet_name=sheet_name)
        except:
            history_df = pd.DataFrame()
    else:
        history_df = pd.DataFrame()

    # Handling new cases
    if not history_df.empty:
        is_new = ~potential_duplicates[['Invoice Number 1', 'Invoice Number 2']].apply(tuple, axis=1).isin(
            history_df[['Invoice Number 1', 'Invoice Number 2']].apply(tuple, axis=1)
        )
        new_duplicates = potential_duplicates[is_new]
        updated_history = pd.concat([history_df, new_duplicates], ignore_index=True)
    else:
        new_duplicates = potential_duplicates
        updated_history = potential_duplicates

    # Update of the historical data
    if os.path.exists(history_file_path):
        with pd.ExcelWriter(
            history_file_path,
            engine='openpyxl',
            mode='a',
            if_sheet_exists='replace'
        ) as writer:
            updated_history.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(
            history_file_path,
            engine='openpyxl',
            mode='w'
        ) as writer:
            updated_history.to_excel(writer, sheet_name=sheet_name, index=False)

    new_duplicates.to_excel(new_file_path, index=False)

    return new_duplicates


# Execution of Check 1: ignoring LE
new_LE = save_new_and_history(
    potential_duplicates=potential_duplicates_le,
    sheet_name="ignoring LE",
    new_file_path="potential_duplicates_no_LE.xlsx",
    history_file_path="potential_duplicates_history.xlsx"
)

# Execution of Check 2: ignoring supplier name
new_supplier = save_new_and_history(
    potential_duplicates=potential_duplicates_supplier,
    sheet_name="ignoring supplier name",
    new_file_path="potential_duplicates_no_supplier.xlsx",
    history_file_path="potential_duplicates_history.xlsx"
)

# LOG
print(f"ignoring LE: {len(new_LE)} new records")
print(f"ignoring supplier name: {len(new_supplier)} new records")