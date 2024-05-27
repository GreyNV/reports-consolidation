import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime



def load_and_combine_entrata_files(source_folder_path, report_type):
    """
    Load and combine files from the specified folder path for a given report type.
    
    Args:
    source_folder_path (str): Path to the folder containing Excel files to be processed.
    report_type (str): Type of report being processed.
    
    Returns:
    pd.DataFrame: Combined DataFrame containing data from all files.
    """
    files = [f for f in os.listdir(source_folder_path) if f.endswith('.xlsx')]
    combined_data = pd.DataFrame()

    for file in files:
        file_path = os.path.join(source_folder_path, file)
        print(f"Processing file: {file_path}")
        workbook = load_workbook(filename=file_path, data_only=True)
        sheet = workbook.active
        data = sheet.values
        columns = next(data)[0:]
        df = pd.read_excel(file_path, header=6)
        print(f"Loaded {len(df)} rows from {file_path}")
        df['Account'] = df['Account'].apply(lambda x: f'E{x}' if pd.notnull(x) else x)
        df = df.dropna(subset=[df.columns[3]])  # Assuming the 4th column needs to have non-null values
        df = df.dropna(subset=[df.columns[0]])  # Assuming the 1th column needs to have non-null values
        df['ReportType'] = report_type
        report_date = sheet['A5'].value  # Assuming the date is in cell A5
        df['ReportDate'] = report_date
        if combined_data.empty:
            combined_data = df
        else:
            combined_data = pd.concat([combined_data, df], ignore_index=True)

    return combined_data

def process_report(software, report_type = ""):
    """
    Process report data for a specific report type.
    
    Args:
    report_type (str): Type of report being processed.
    
    Returns:
    pd.DataFrame: Processed report data.
    """
    ReportPath = f"c://Users/AndriiRybak/OneDrive - EVU Residential/Asset Management BI project/Source files/{software}/{report_type}/"
    if software == "Entrata":
        report_data = load_and_combine_entrata_files(ReportPath, report_type)

    elif software == "Yardi":
        report_data = load_and_combine_yardi_files(ReportPath, report_type)

    elif software == "Resman":
        report_data = load_and_combine_resman_files(ReportPath)
    
    return report_data

def load_and_combine_yardi_files(source_folder_path, report_type):
    """
    Load and combine files from the specified folder path for a given report type.
    
    Args:
    source_folder_path (str): Path to the folder containing Excel files to be processed.
    report_type (str): Type of report being processed.
    
    Returns:
    pd.DataFrame: Combined DataFrame containing data from all files.
    """
    files = [f for f in os.listdir(source_folder_path) if f.endswith('.xlsx')]
    combined_data = pd.DataFrame()

    for file in files:
        file_path = os.path.join(source_folder_path, file)
        print(f"Processing file: {file_path}")
        
        df = pd.read_excel(file_path, nrows=3, header=None)
        print(f"Extracting property name and date from {file_path}")
        property_name = df.iloc[0, 0]
        property_name = property_name.split(' (')[0]
        print(f"Property name: {property_name}")
        try:
            report_period_str = df.iloc[2, 0]  # Adjusted to third row (index 2)
            report_date_str = report_period_str.split('=')[1].strip()
            # Convert the report date string to a datetime object for standardized formatting
            report_date = datetime.strptime(report_date_str, '%b %Y')
        except (IndexError, ValueError):
            # In case of any unexpected format, set a default or indicate an error
            report_date = "Date extraction failed"  # or use datetime.now() or another placeholder
        print(f"Report date: {report_date}")

        df = pd.read_excel(file_path, header=5)
        df.rename(columns={df.columns[0]: 'Account', df.columns[1]: 'Account Name'}, inplace=True)
        print(f"Loaded {len(df)} rows from {file_path}")
        df['Account'] = df['Account'].apply(lambda x: f'Y{x}' if pd.notnull(x) else x)

        df = df.dropna(subset=[df.columns[0]])  # Assuming the 4th column needs to have non-null values
        df['ReportType'] = report_type
        

        df['ReportDate'] = report_date
        df['Property'] = property_name
        df["Amount"] = df['Credit'] - df['Debit']
        if combined_data.empty:
            combined_data = df
        else:
            combined_data = pd.concat([combined_data, df], ignore_index=True)

    return combined_data

def load_and_combine_resman_files(source_folder_path):
    """
    Load and combine files from the specified folder path for a given report type.
    
    Args:
    source_folder_path (str): Path to the folder containing Excel files to be processed.
    report_type (str): Type of report being processed.
    
    Returns:
    pd.DataFrame: Combined DataFrame containing data from all files.
    """
    files = [f for f in os.listdir(source_folder_path) if f.endswith('.csv')]
    combined_data = pd.DataFrame()

    for file in files:

        file_path = os.path.join(source_folder_path, file)
        print(f"Processing file: {file_path}")
        
        try :
            df = pd.read_csv(file_path)
        except pd.errors.EmptyDataError:
            continue
        print(f"Extracting property name and date from {file_path}")
        
        property_name = file.split('_')[1]
        print(f"Property name: {property_name}")

        report_date = file.split('_')[0]
        print(f"Report date: {report_date}")

        report_type = file.split('_')[2].split('.')[0]
        print(f"Report type: {report_type}")

        df.rename(columns={df.columns[0]: 'Account', df.columns[1]: 'Account Name', df.columns[6]: 'Debit', df.columns[7]: 'Credit'}, inplace=True)
        print(f"Loaded {len(df)} rows from {file_path}")
        
        df['Account'] = df['Account'].apply(lambda x: f'R{x}' if pd.notnull(x) else x)
        df['Credit'] = df['Credit'].apply(lambda x: x if pd.notnull(x) else 0)
        df['Debit'] = df['Debit'].apply(lambda x: x if pd.notnull(x) else 0)

        df = df.dropna(subset=[df.columns[3]])  # Assuming the 4th column needs to have non-null values
        df['ReportType'] = report_type
        df['ReportDate'] = report_date
        df['Property'] = property_name
        df["Amount"] = df['Credit'] - df['Debit']
        if combined_data.empty:
            combined_data = df
        else:
            combined_data = pd.concat([combined_data, df], ignore_index=True)

    return combined_data
# Process files for different report types
entrata_cash = process_report(report_type="Cash", software="Entrata")
yardi_cash = process_report(report_type="Cash", software="Yardi")

resman_data = process_report(software = "Resman")

# Removed accrual books, plan to create a separate report for accrual books to not exceed the excel row limit for a sheet
# accrual_books_data = process_report("Accrual")


# Debugging information
print("Successfully loaded:")
print(entrata_cash.head())
print(yardi_cash.head())
print(resman_data.head())
# print("\nNet Change Data:")
# print(accrual_books_data.head())

# print("\nAppending information from Entrata Reports")
# Append data from different report types
# Replaced appended data with just cash books data
# appended_data = pd.concat([entrata_cash, accrual_books_data], ignore_index=True)

print("\nUnpivoting properties data")
# Specify the columns that you want to keep fixed. For example, 'Account', 'Account Name', 'ReportDate', 'ReportType'
fixed_columns = ['Account', 'Account Name', 'ReportDate', 'ReportType']

# Specify the columns that you want to unpivot or melt. In this case, it would be the columns that are not in 'fixed_columns'
value_vars = [column for column in entrata_cash.columns if column not in fixed_columns]

# Use the melt function to unpivot the DataFrame
unpivoted_data = pd.melt(entrata_cash, id_vars=fixed_columns, value_vars=value_vars, var_name='Property', value_name='Amount')

# converting data to negative values to have the income as positive as expense as negative number
unpivoted_data['Amount'] *= -1


# Now 'unpivoted_data' contains the unpivoted version of your DataFrame. Let's remove the rows where the amount is 0
filtered_data = unpivoted_data.query('Amount != 0')
print("Successfully loaded:")
print(filtered_data.head())
print(yardi_cash.head())
print(resman_data.head())
print("Saving data into excel")
# Save the resulting DataFrame to an Excel file
filtered_data.to_excel("c://Users/AndriiRybak/OneDrive - EVU Residential/Asset Management BI project/Consolidation - Entrata.xlsx", index=False)
yardi_cash.to_excel("c://Users/AndriiRybak/OneDrive - EVU Residential/Asset Management BI project/Consolidation - Yardi.xlsx", index=False)
print("Done! Exiting code...")
resman_data.to_excel("c://Users/AndriiRybak/OneDrive - EVU Residential/Asset Management BI project/Consolidation - Resman.xlsx", index=False)
