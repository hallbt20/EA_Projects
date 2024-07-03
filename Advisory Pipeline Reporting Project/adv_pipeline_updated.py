import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from re import sub
import warnings


# Global vars
selected_reports = []


def find_file(folder_path, file_start):
    for file_name in os.listdir(folder_path):
        if file_name.startswith(file_start):
            return os.path.join(folder_path, file_name)
    return None


def browse_for_file(missing_files):
    for key, value in missing_files.items():
        if not value:
            file_path = filedialog.askopenfilename(title=f"Locate {key}")
            if file_path:
                missing_files[key] = file_path
            else:
                messagebox.showwarning("File not found", f"{key} file not found. Please locate the file.")
    return missing_files


def start_window_adv():
    folder_path = filedialog.askdirectory(title="Select a folder")
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
        return None

    files_to_find = {
        "sf_active_file": "Salesforce Active",
        "sf_closed_file": "Salesforce Closed",
        "ns_active_file": "Netsuite Active",
        "ns_closed_file": "Netsuite Closed",
        "great_lakes_file": "EAG GL",
        'originators_list': 'Originators List',
    }

    found_files = {key: find_file(folder_path, start) for key, start in files_to_find.items()}

    missing_files = {key: value for key, value in found_files.items() if value is None}

    if missing_files:
        messagebox.showinfo("Missing Files", "Some files were not found. Please locate the missing files.")
        found_files.update(browse_for_file(missing_files))

    file_path_vars = [
        found_files['sf_active_file'],
        found_files['sf_closed_file'],
        found_files['ns_active_file'],
        found_files['ns_closed_file'],
        found_files['great_lakes_file'],
        found_files['originators_list'],
    ]

    return file_path_vars


def start_window():
    folder_path = filedialog.askdirectory(title="Select a folder")
    if not folder_path:
        messagebox.showerror("Error", "No folder selected.")
        return None

    files_to_find = {
        "adv_out_file": "ADV OUT Pipeline",
        "sf_active_file": "Salesforce Active",
        "sf_closed_file": "Salesforce Closed",
        "ns_active_file": "Netsuite Active",
        "ns_closed_file": "Netsuite Closed",
        "eag_gc_oit_file": "EAG GC OIT",
        "hubspot_file": "hubspot",
        "great_lakes_file": "EAG GL",
        'originators_list': 'Originators List'
    }

    found_files = {key: find_file(folder_path, start) for key, start in files_to_find.items()}

    missing_files = {key: value for key, value in found_files.items() if value is None}

    if missing_files:
        messagebox.showinfo("Missing Files", "Some files were not found. Please locate the missing files.")
        found_files.update(browse_for_file(missing_files))

    file_path_vars = [
        found_files['sf_active_file'],
        found_files['sf_closed_file'],
        found_files['ns_active_file'],
        found_files['ns_closed_file'],
        found_files['great_lakes_file'],
        found_files['originators_list'],
        found_files['adv_out_file'],
        found_files['eag_gc_oit_file'],
        found_files['hubspot_file']
    ]

    return file_path_vars


def startup_window():
    def on_browse():
        if advisory_var.get():
            selected_reports.append('Advisory')
        if outsourcing_var.get():
            selected_reports.append('Outsourcing')
        if not selected_reports:
            messagebox.showwarning("Selection Error", "Please select at least one report type.")
            return

        root.destroy()

    root = tk.Tk()
    root.title("Advisory Pipeline Automatic Report Generator")

    label = tk.Label(root, text="This is the Advisory Pipeline Automatic Report Generator.", font=('Helvetica', 14, 'bold'))
    label.pack(pady=(20, 10))

    instruction = tk.Label(root, text="Please select a folder with the raw data files.", font=('Helvetica', 12))
    instruction.pack(pady=(0, 20))

    advisory_var = tk.BooleanVar()
    outsourcing_var = tk.BooleanVar()

    advisory_check = tk.Checkbutton(root, text="Advisory Report", variable=advisory_var, font=('Helvetica', 12))
    advisory_check.pack(pady=(0, 10))

    outsourcing_check = tk.Checkbutton(root, text="Outsourcing Report", variable=outsourcing_var, font=('Helvetica', 12))
    outsourcing_check.pack(pady=(0, 10))

    browse_button = tk.Button(root, text="Browse", command=on_browse, font=('Helvetica', 12))
    browse_button.pack(pady=(20, 20))

    root.mainloop()

    if 'Advisory' in selected_reports:
        return start_window_adv()
    else:
        return start_window()


def dataframe_ordering(df):
    df = df.rename(columns={
        'Opp Number': 'Opportunity ID',
        'Service Line Group (EA)': 'Service Line Group',
        'Service Line L1': 'Service Line Group',
        'Service Status': 'Stage',
        'Account Name: Account Name': 'Account Name',
        'Organization': 'Account Name',
        'First Year Fees (EA\'s portion)': 'First Year Fees',
        'Estimated Fee': 'First Year Fees',
        'Total Contract Value (EA\'s portion)': 'Total Contract Value',
        'Create Date': 'Created Date',
        'Service Status Change Date': 'Close Date',
        'Days Open': 'Age',
        'Originator': 'Opportunity Originator',
        'Opp Leader': 'Opportunity Leader',
        'Other Contributors': 'Opportunity Team',
        'Service': 'Service Lines',
        'Opportunity Type': 'Type',
        'Account Name: Industry': 'Industry',
        'Office Location Client Assigned to': 'Office Location',
        'Office': 'Office Location',
        'Recurring or One Time?': 'Recurrence'
    })

    df = df[[
        'Data Source',
        'Opportunity ID',
        'Service Line Group',
        'Stage',
        'Stage (adjusted)',
        'Account Name',
        'Opportunity Name',
        'First Year Fees',
        'Total Contract Value',
        'Created Date',
        'Close Date',
        'Age',
        'Opportunity Originator',
        'Opportunity Leader',
        'Opportunity Team',
        'Service Lines',
        'Type',
        'Industry',
        'Office Location',
        'Recurrence',
        'Contract Duration',
        'Last Activity',
        'Next Step',
        'Next Step Due Date',
        'Client Code',
        'Primary Campaign Source: Campaign Name',
        'Originator Service Line',
        'Opp Leader Service Line',
        'Contact'
    ]]

    return df


def clean_salesforce(df):
    """
    Function that cleans and formats both active and closed Salesforce data
    :param df: Dataframe containing active or closed records from Salesforce
    :return: Updated dataframe with added columns, duplicates removed, and proper formatting
    """
    # Add columns to dataframe
    df['Data Source'] = 'Salesforce'
    df['Stage (adjusted)'] = df['Stage']

    # Initialize empty columns
    for col in ['Originator Service Line', 'Opp Leader Service Line', 'Contact']:
        df[col] = None

    # Remove duplicates and keep only the first occurrence for 'Opportunity ID' (this is case-sensitive)
    df = df.drop_duplicates(subset='Opportunity ID', keep='first')

    return dataframe_ordering(df)


def clean_netsuite(df, active=True):
    # Add columns 'Data source' to dataframes
    df['Data Source'] = 'NetSuite'

    # Initialize empty columns
    for col in [
        'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date', 'Client Code',
        'Primary Campaign Source: Campaign Name'
    ]:
        df[col] = None

    # Fill empty 'Service Description' values with corresponding value from 'Service'
    df['Opportunity Name'] = df['Service Description'].fillna(df['Service'])

    # Create 'Stage (adjusted)' column based on 'Service Status' values
    if active:
        df['Stage (adjusted)'] = df['Service Status'].apply(
            lambda x: 'Qualified' if x == 'Prospect - Active' else 'Proposal'
        )
    else:
        df['Stage (adjusted)'] = df['Service Status'].apply(
            lambda x: 'Closed Won' if x == 'Won' else 'Closed Lost'
        )

    df['Total Contract Value'] = df['Estimated Fee']

    return dataframe_ordering(df)


def find_sheet_name(file_path, sheet_name_start):
    sheet_names = pd.ExcelFile(file_path).sheet_names
    for sheet_name in sheet_names:
        if sheet_name.startswith(sheet_name_start):
            return sheet_name
    return None


def clean_great_lakes(df, active=True):
    # Columns to add
    df['Data Source'] = 'Great Lakes'
    df['Service Line Group (EA)'] = 'ADV'
    df['Stage'] = df['Stage (adjusted)']
    df['Opportunity Leader'] = 'Marc Moskowitz'
    df['Service Lines'] = 'ADV Workday Adaptive Planning'

    # Columns to initialize
    cols_to_initialize = [
        'Opportunity Team', 'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date', 'Contact',
        'Client Code', 'Primary Campaign Source: Campaign Name', 'Originator Service Line', 'Opp Leader Service Line',
    ]

    for col in cols_to_initialize:
        df[col] = None

    df['Opportunity Name'] = df['Opportunity Name'].apply(
        lambda x: 'Workday Adaptive Planning - Implementation' if str(x).replace(' ', '').startswith('W-I') else
        'Workday Adaptive Planning - Optimization' if str(x).replace(' ', '').startswith('W-O') else
        'Workday Adaptive Planning - Add On' if str(x).replace(' ', '').startswith('W-AddOn') else
        'Workday Adaptive Planning - Prepaid Block' if str(x).replace(' ', '').startswith('W-B') else
        'Workday Adaptive Planning - TM' if str(x).replace(' ', '').startswith('W-TM') else
        'Workday Adaptive Planning - Renewal' if str(x).replace(' ', '').startswith('W-Renewal') else
        'Workday Adaptive Planning - Advisory' if str(x).startswith('Advisory') or str(x).endswith('Advisory') else
        'Workday Adaptive Planning - MS' if str(x).replace(' ', '').startswith('W-MS') else
        None
    )

    # Populate empty values in 'First Year Fees' with corresponding values from 'Total Contract Value'
    df['First Year Fees (EA\'s portion)'] = df['First Year Fees (EA\'s portion)'].fillna(df['Total Contract Value (EA\'s portion)'])

    if active:
        df = df[~df['Stage (adjusted)'].isin(['Closed Lost', 'Closed Won', '']) & ~df['Stage (adjusted)'].isna()]

        df['Stage (adjusted)'] = df['Stage'].apply(
            lambda x: 'Qualified' if x in ['AOTP', 'Discovery', 'Scoping', 'SQL'] else
            'Unqualified' if x in ['Cold SQL', 'MQL'] else
            'Proposal' if x in ['Renewals', 'SOW', 'SOW - No Decision', 'Verbal Approval'] else
            'None'
        )
    else:
        df = df[~df['Close Date'].isna() & (df['Close Date'] >= pd.Timestamp('2023-08-01'))]

    df = df[df['Opportunity Originator'] != 'Old Old']

    return dataframe_ordering(df)


def prompt_adv_values(unique_originators):
    root = tk.Tk()
    root.title("Fill in ADV? values")

    entries = []
    new_entries = {}

    def check_values():
        all_selected = all(entry[1].get() != "Select" for entry in entries)
        submit_button.config(state=tk.NORMAL if all_selected else tk.DISABLED)

    def on_submit():
        for entry in entries:
            unique_originators[entry[0]] = entry[1].get()
        root.quit()
        root.destroy()

    # Create labels and dropdowns for each unique unmatched originator
    for i, originator in enumerate(unique_originators):
        label = tk.Label(root, text=f"{originator}")
        label.grid(row=i, column=0)
        var = tk.StringVar(root)
        var.set("Select")
        dropdown = ttk.Combobox(root, textvariable=var, values=["ADV", "Other"])
        dropdown.grid(row=i, column=1)
        dropdown.bind("<<ComboboxSelected>>", lambda event: check_values())
        entries.append((originator, var))

    # Create a submit button, initially disabled
    submit_button = tk.Button(root, text="Submit", command=on_submit, state=tk.DISABLED)
    submit_button.grid(row=len(unique_originators), column=0, columnspan=2)

    # Run the Tkinter main loop
    root.mainloop()

    return unique_originators


def clean_combined_adv(df, originator_list=None, active=True):
    df = df[df['Service Line Group'].isin(['ADV', 'ADV (Advisory)'])]

    df.loc[:, 'Type'] = df.apply(
        lambda row: 'Renewal Business' if row['Type'] in [
            'Existing Business', 'Renewal Business', 'Renewal of Existing Business'
        ]
        else 'Expanded Business' if row['Type'] in [
            'Expanded Business', 'Expansion of Existing Services', 'New Service for Existing Client',
            'New Service(s) for Existing Client'
        ]
        else 'New Business' if row['Type'] in ['NEW', 'New', 'New Business', 'New Client']
        else 'Renewal Business' if row['Type'] == 'TBD' and row['Service Lines'] == 'Maintenance Renewal'
        else 'Expanded Business' if row['Type'] == 'TBD' and row['Service Lines'] != 'Maintenance Renewal'
        else row['Type'],
        axis=1
    )

    # Fill empty 'Opportunity Originator' values with corresponding value from 'Opportunity Leader'
    df.loc[:, 'Opportunity Originator'] = df['Opportunity Originator'].fillna(df['Opportunity Leader'])

    if not active:
        originator_df = pd.read_excel(originator_list)

        df.loc[:, 'Opportunity Originator'] = df.loc[:, 'Opportunity Originator'].apply(
            lambda x: sub(r'\s*\(.*?\)$', '', x)
        )

        # Convert both columns to sets
        df_set = set(df['Opportunity Originator'])
        originator_set = set(originator_df['Opportunity Originator'])

        # Find the difference between the two sets
        diff_dict = {key: None for key in df_set - originator_set}

        if len(diff_dict):
            # Prompt user to determine the 'ADV?' value for each unique unmatched originator
            new_entries = prompt_adv_values(diff_dict)

            # Convert the dictionary to a dataframe
            new_entries_df = pd.DataFrame(list(new_entries.items()), columns=['Opportunity Originator', 'ADV?'])
            new_entries_df['Originator Service Line'] = None

            # Concatenate the original dataframe with the new entries dataframe
            originator_df = pd.concat([originator_df, new_entries_df], ignore_index=True)

            originator_df = originator_df.sort_values(by='Opportunity Originator').reset_index(drop=True)

            try:
                originator_df.to_excel(originator_list, index=False)
            except PermissionError:
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showerror(
                    "Permission Error",
                    "The file 'Originators List.xlsx' appears to be open. Please close the file and press OK."
                )
                root.destroy()
                originator_df.to_excel(originator_list, index=False)

        # Merge the two dataframes on the 'Opportunity Originator' column
        merged_df = pd.merge(
            df,
            originator_df[['Opportunity Originator', 'ADV?']],
            on='Opportunity Originator',
            how='left'
        )

        # Get the list of columns
        cols = list(merged_df.columns)

        # Move 'Originator in ADV?' behind 'Opportunity Originator'
        originator_col = cols.pop(cols.index('ADV?'))
        cols.insert(cols.index('Opportunity Originator') + 1, originator_col)

        df = merged_df[cols]

        df.loc[:, 'Stage (adjusted)'] = pd.Categorical(
            df['Stage (adjusted)'],
            categories=['Closed Won', 'Closed Lost'],
            ordered=True
        )

        df = df.sort_values(by=['Stage (adjusted)', 'Close Date'], ascending=[False, False])

        return df

        # Convert 'Stage (adjusted)' to a categorical type with the specified order

    df.loc[:, 'Stage (adjusted)'] = pd.Categorical(
        df['Stage (adjusted)'],
        categories=['Proposal', 'Qualified', 'Unqualified', 'Suspect'],
        ordered=True
    )

    # Sort the dataframe
    df = df.sort_values(by=['Stage (adjusted)', 'Created Date'], ascending=[True, False])

    return df


def show_report_generated_message():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo(
        "Report Generated",
        "Your report has been generated. It is located in your Documents folder, and it is called ADV Pipeline.xlsx"
    )
    root.destroy()


if __name__ == "__main__":
    # Starts the UI window for selecting a folder with contained files
    report_files = startup_window()

    # Read, cleans, and reformats Salesforce data to a Pandas dataframe
    sf_active = clean_salesforce(pd.read_excel(report_files[0]))
    sf_closed = clean_salesforce(pd.read_excel(report_files[1]))

    # Read, clean, and NetSuite data as Pandas dataframes and returns cleaned data frame
    ns_active = clean_netsuite(pd.read_excel(report_files[2]))
    ns_closed = clean_netsuite(pd.read_excel(report_files[3]), False)

    great_lakes_file = report_files[4]
    great_lakes_active_sheet = find_sheet_name(great_lakes_file, 'ADV Active')
    great_lakes_closed_sheet = find_sheet_name(great_lakes_file, 'ADV Closed')

    great_lakes_active = clean_great_lakes(pd.read_excel(great_lakes_file, sheet_name=great_lakes_active_sheet))
    great_lakes_closed = clean_great_lakes(pd.read_excel(great_lakes_file, sheet_name=great_lakes_closed_sheet), False)

    if 'Advisory' in selected_reports:
        # ADV Report start
        with warnings.catch_warnings(action='ignore'):
            all_active = pd.concat([sf_active, ns_active, great_lakes_active], ignore_index=True)
            all_closed = pd.concat([sf_closed, ns_closed, great_lakes_closed], ignore_index=True)

        # Split off Advisory data
        adv_active = clean_combined_adv(all_active)
        adv_closed = clean_combined_adv(all_closed, report_files[5], False)

        # Create an ExcelWriter object and specify the file path
        with pd.ExcelWriter('ADV Pipeline.xlsx', engine='xlsxwriter') as writer:
            # Write each dataframe to a different sheet

            adv_active.to_excel(writer, sheet_name='ADV Active', index=False)
            adv_closed.to_excel(writer, sheet_name='ADV Closed', index=False)
            pd.read_excel(report_files[5]).to_excel(writer, sheet_name='Originators List', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Originators List']
            worksheet.hide()

        # Call the function at the end of your script
        show_report_generated_message()
    else:
        print("This is all I got for now.")
