import pandas as pd


def process_strings(string_list):
    processed_list = []

    for string in string_list:
        # Remove the '[C]' suffix if it exists
        if string.endswith(' [C]'):
            string = string[:-4]

        # Split the string into words
        words = string.split()

        # Switch the first and last words
        if len(words) > 1:
            words[0], words[-1] = words[-1], words[0]

        # Join the words back into a single string
        processed_string = ' '.join(words)

        # Add the processed string to the list
        processed_list.append(processed_string)

    return processed_list


rmd_file = 'Employee Summary (RMD - By Month).xlsx'

# Load the Excel file (Excel file contains monthly advisor information sheet-by-sheet)
excel_file = pd.ExcelFile(rmd_file)

# Store sheet names as sheet_names
sheet_names = excel_file.sheet_names

# Create a dictionary to store the month-by-month DataFrames
rmd_dataframes = {
    sheet: pd.read_excel(rmd_file, sheet_name=sheet, header=2) for sheet in sheet_names
}

# Add an extra column for WIP Month
month_counter = 8

for sheet in sheet_names:
    if not month_counter % 12:
        rmd_dataframes[sheet]['WIP_Month'] = 12
    else:
        rmd_dataframes[sheet]['WIP_Month'] = month_counter % 12

    month_counter += 1

# Concatenate all the DataFrames into one overall DataFrame
fiscal_year_by_month_rmd = pd.concat(rmd_dataframes.values(), ignore_index=True)

fiscal_year_by_month_rmd = fiscal_year_by_month_rmd[
    ['ID', 'Employee Name', 'Employee Location', 'Emp Title', 'Emp Service1', 'Emp Service2', 'Emp Service3',
     'WIP_Month', 'Actual Billable Hrs']]

# Rename the columns
fiscal_year_by_month_rmd.rename(columns={
    'ID': 'ID',
    'Employee Name': 'Employee Name',
    'Employee Location': 'Office Location',
    'Emp Title': 'Title',
    'Emp Service1': 'Branch',
    'Emp Service2': 'Practice Group',
    'Emp Service3': 'Practice',
    'Actual Billable Hrs': 'WIP_Hours_RMD'
}, inplace=True)

# Apply a similar process for ETD Data
etd_file = 'Employee Summary (ETD - By Month).xlsx'

# Load the Excel file (Excel file contains monthly advisor information sheet-by-sheet)
excel_file = pd.ExcelFile(etd_file)

# Store sheet names as sheet_names
sheet_names = excel_file.sheet_names

# Create a dictionary to store the month-by-month DataFrames
etd_dataframes = {
    sheet: pd.read_excel(etd_file, sheet_name=sheet) for sheet in sheet_names
}

# Add an extra column for WIP Month
month_counter = 8

for sheet in sheet_names:
    if not month_counter % 12:
        etd_dataframes[sheet]['WIP_Month'] = 12
    else:
        etd_dataframes[sheet]['WIP_Month'] = month_counter % 12

    month_counter += 1

# Concatenate all the DataFrames into one overall DataFrame
fiscal_year_by_month_etd = pd.concat(etd_dataframes.values(), ignore_index=True)

fiscal_year_by_month_etd = fiscal_year_by_month_etd[[
    'FullName', 'Billable Hours', 'WIP_Month'
]]

# Rename the columns
fiscal_year_by_month_etd.rename(columns={
    'Billable Hours': 'WIP_Hours_ETD'
}, inplace=True)

# Apply the lambda function to switch first and last names
fiscal_year_by_month_etd['Employee Name'] = fiscal_year_by_month_etd['FullName'].apply(lambda name: ' '.join(name.split(' ')[::-1]))

merged_df = pd.merge(fiscal_year_by_month_etd, fiscal_year_by_month_rmd, how='outer', on=['Employee Name', 'WIP_Month'])

month_map = {
    1: 'January', 2: 'February', 3: 'March', 4: 'April',
    5: 'May', 6: 'June', 7: 'July', 8: 'August',
    9: 'September', 10: 'October', 11: 'November', 12: 'December'
}

# Create a new column WIP_Month_Name using the map
merged_df['WIP_Month_Name'] = merged_df['WIP_Month'].map(month_map)

merged_df = merged_df[[
    'ID', 'Employee Name', 'Office Location', 'Title', 'Branch', 'Practice Group', 'Practice', 'WIP_Month',
    'WIP_Month_Name', 'WIP_Hours_RMD', 'WIP_Hours_ETD'
]]

merged_df['Variance'] = merged_df['WIP_Hours_RMD'] - merged_df['WIP_Hours_ETD']


# Define the function to determine variance type
def determine_variance_type(variance, wip_hours_rmd, wip_hours_etd):
    if pd.isna(variance):
        if (pd.isna(wip_hours_rmd) or wip_hours_rmd == 0) and (pd.isna(wip_hours_etd) or wip_hours_etd == 0):
            return 'Missing - Good'
        else:
            return 'Missing - Bad'
    elif abs(variance) < 0.5:
        return 'Good'
    else:
        return 'Bad'


# Apply the function to create the 'Variance_Type' column
merged_df['Variance_Type'] = merged_df.apply(
    lambda row: determine_variance_type(row['Variance'], row['WIP_Hours_RMD'], row['WIP_Hours_ETD']), axis=1
)

cont_workers = pd.read_excel('Advisory Contractors.xlsx')

cont_workers_list = cont_workers['Contingent Worker'].to_list()

cont_workers_list = process_strings(cont_workers_list)

merged_df.loc[merged_df['Employee Name'].isin(cont_workers_list), 'Title'] = 'Contingent Worker'


# Function to save DataFrame to Excel with formatting and filters
def save_to_excel_with_formatting_and_filters(df, file_name):
    # Save DataFrame to Excel without formatting first
    df.to_excel(file_name, index=False, engine='openpyxl')

    # Load the workbook and select the active worksheet
    from openpyxl import load_workbook
    workbook = load_workbook(file_name)
    worksheet = workbook.active

    # Add filters to the top row
    worksheet.auto_filter.ref = worksheet.dimensions

    # Set the column widths based on the maximum length in each column
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    # Save the workbook with the adjusted column widths and filters
    workbook.save(file_name)


# Save the DataFrame to Excel with formatting and filters
save_to_excel_with_formatting_and_filters(merged_df, 'merged.xlsx')

pass