from datetime import datetime
import pandas as pd
from re import sub

from popups import *


def find_sheet_name(file_path, sheet_name_start):
    sheet_names = pd.ExcelFile(file_path).sheet_names
    for sheet_name in sheet_names:
        if sheet_name.startswith(sheet_name_start):
            return sheet_name
    return None


def dataframe_ordering(df):
    df = df.rename(columns={
        'Data source': 'Data Source',
        'Opp Number': 'Opportunity ID',
        'QUOTE NUMBER': 'Opportunity ID',
        'Opp ID': 'Opportunity ID',
        'Opp Name': 'Opportunity Name',
        'Service Line Group (EA)': 'Service Line Group',
        'Service Line L1': 'Service Line Group',
        'Service Status': 'Stage',
        'STAGE': 'Stage',
        'Deal Stage': 'Stage',
        'Stage adjusted': 'Stage (adjusted)',
        'Account Name: Account Name': 'Account Name',
        'Organization': 'Account Name',
        'ACCOUNT_NAME': 'Account Name',
        'Account name': 'Account Name',
        'OPPORTUNITY_NAME': 'Opportunity Name',
        'First Year Fees (EA\'s portion)': 'First Year Fees',
        'Estimated Fee': 'First Year Fees',
        'ESTIMATED_FEES': 'First Year Fees',
        'SOW and Commission': 'Total Contract Value',
        'Total Contract Value (EA\'s portion)': 'Total Contract Value',
        'Create Date': 'Created Date',
        'CREATED_DATE': 'Created Date',
        'EXPECTED_CLOSE': 'Close Date',
        'Service Status Change Date': 'Close Date',
        'Closed Date': 'Close Date',
        'Days Open': 'Age',
        'Originator': 'Opportunity Originator',
        'ORIGINATOR': 'Opportunity Originator',
        'Opp Leader': 'Opportunity Leader',
        'SALES_LEADER': 'Opportunity Leader',
        'Leader': 'Opportunity Leader',
        'Other Contributors': 'Opportunity Team',
        'Team': 'Opportunity Team',
        'Service': 'Service Lines',
        'SERVICE_LINE': 'Service Lines',
        'Service Line': 'Service Lines',
        'Opportunity Type': 'Type',
        'Account Name: Industry': 'Industry',
        'INDUSTRY': 'Industry',
        'Industry/Segment': 'Industry',
        'Industry Group': 'Industry',
        'Office Location Client Assigned to': 'Office Location',
        'Office': 'Office Location',
        'Recurring or One Time?': 'Recurrence',

        # Hubspot
        'Record ID': 'Opportunity ID',
        'Deal Name': 'Account Name',
        'Service Team': 'Opportunity Name',
        'Amount': 'First Year Fees',
        'Amount in company currency': 'Total Contract Value',
        'Deal owner': 'Opportunity Leader',
        'Source/Referral': 'Opportunity Team'
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


def clean_great_lakes(df):
    # Columns to add
    df['Data Source'] = 'Great Lakes'
    df['Service Line Group'] = 'ADV'
    df['Opportunity Leader'] = 'Marc Moskowitz'
    df['Service Lines'] = 'ADV Workday Adaptive Planning'

    # Columns to initialize
    cols_to_initialize = [
        'Opportunity Team', 'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date', 'Contact',
        'Client Code', 'Primary Campaign Source: Campaign Name', 'Originator Service Line', 'Opp Leader Service Line',
        'Opportunity ID', 'Office Location', 'Recurrence'
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
    df['Amount'] = df['Amount'].fillna(df['SOW and Commission'])

    # Replace 'Old Old' with 'Marc Moskowitz' in the 'Opportunity Originator' column
    df['Opportunity Originator'] = df['Created By'].replace('Old Old', 'Marc Moskowitz')

    df['Stage (adjusted)'] = df['Stage'].apply(
        lambda x: 'Qualified' if x in ['AOTP', 'Discovery', 'Scoping', 'SQL'] else
        'Unqualified' if x in ['Cold SQL', 'MQL'] else
        'Proposal' if x in ['Renewals', 'SOW', 'SOW - No Decision', 'Verbal Approval'] else
        x
    )

    return dataframe_ordering(df)


def clean_hubspot(df):
    df['Data Source'] = 'Hubspot'
    df['Service Line Group'] = 'OUT'
    df['Type'] = df['Deal Type']
    df['Service'] = df['Service Team']

    df['Close Date'] = pd.to_datetime(df['Close Date'])

    cols_to_initialize = [
        'Created Date', 'Age', 'Opportunity Originator', 'Office Location Client Assigned to',
        'Recurrence', 'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date', 'Client Code',
        'Primary Campaign Source: Campaign Name', 'Originator Service Line', 'Opp Leader Service Line', 'Contact'
    ]

    for col in cols_to_initialize:
        df[col] = None

    # Filter dataframe where Deal Stage is only one of the following
    df = df[df['Deal Stage'].isin([
        'Inquiry',
        'Intro Call Scheduled',
        'BD Action Required',
        'Consideration / Materials Sent',
        'EA Incoming Leads',
        'Assessment',
        'Engagement Letter Sent',
        'Dead Leads Inquiries',
        'Closed lost'
    ])]

    df = df[df['EA Opportunity'].isna()].drop(columns=['EA Opportunity'])

    df = df[df['Service Team'] == 'Startup']

    df['Stage (adjusted)'] = df['Deal Stage'].apply(
        lambda x: 'Closed Lost' if x in ['Dead Leads Inquiries', 'Closed lost'] else
        'Proposal' if x == 'Engagement Letter Sent' else
        'Qualified' if x in ['Consideration / Materials Sent', 'BD Action Required'] else
        'Unqualified' if x in ['Inquiry', 'Intro Call Scheduled', 'EA Incoming Leads', 'Assessment'] else
        'None'
    )

    df = df[~(df['Deal owner'] == 'John Delalio')]

    return dataframe_ordering(df)


def clean_pnt(df, active=True):
    df = df.copy()
    df['Data Source'] = 'PNT Connectwise'
    df['Service Line Group (EA)'] = 'OUT'

    cols_to_initialize = [
        'Age', 'Opportunity Team', 'Account Name: Industry', 'Office Location Client Assigned to', 'Recurrence',
        'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date',
        'Client Code', 'Primary Campaign Source: Campaign Name', 'Originator Service Line',
        'Opp Leader Service Line', 'Contact'
    ]

    for col in cols_to_initialize:
        df[col] = None

    if active:
        df['Stage (adjusted)'] = df['Sales_Stage']
    else:
        df['Stage (adjusted)'] = df['Sales_Stage'].apply(lambda x: 'Closed Won' if 'Won' else 'Closed Lost')

    df['Total Contract Value (EA\'s portion)'] = df['Total_Revenue'] - df['Product_Cost']
    df.drop(columns=['Total_Revenue', 'Product_Cost'], inplace=True)

    df['First Year Fees (EA\'s portion)'] = df['Total Contract Value (EA\'s portion)']

    df.rename(columns={
        'RecID': 'Opportunity ID',
        'Sales_Stage': 'Stage',
        'Company_Name': 'Account Name: Account Name',
        'Opp_Name': 'Opportunity Name',
        'Created_Date': 'Created Date',
        'Expected_Close_Date': 'Close Date',
        'Originator': 'Opportunity Originator',
        'Opp_Leader': 'Opportunity Leader',
        'Service_Area': 'Service Lines',
        'Service_Type': 'Type',
    }, inplace=True)

    return dataframe_ordering(df)


def clean_legacy(df):
    df['Data Source'] = 'Legacy OIT'
    df['Service Line Group'] = 'OUT'
    df['Total Contract Value'] = df['ESTIMATED_FEES']

    cols_to_initialize = [
        'Office Location', 'Recurrence', 'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date',
        'Client Code', 'Primary Campaign Source: Campaign Name', 'Originator Service Line', 'Opp Leader Service Line',
        'Contact', 'Age', 'Opportunity Team'
    ]

    for col in cols_to_initialize:
        df[col] = None

    df['Stage (adjusted)'] = df['STAGE'].apply(
        lambda x: 'Proposal' if x == 'Quote in Review' else
        'Qualified' if x == 'In Discussion' else
        'Unqualified' if x == 'Lead' else
        x
    )

    df['Type'] = df['TYPE'].apply(
        lambda x: 'New Business' if x == 'NEW' else
        'Expanded Business' if x == 'EXPANDED BUSINESS' else
        None
    )

    return dataframe_ordering(df)


def clean_triangle(df):
    cols_to_initialize = [
        'Recurrence', 'Contract Duration', 'Last Activity', 'Next Step', 'Next Step Due Date', 'Client Code',
        'Primary Campaign Source: Campaign Name', 'Originator Service Line', 'Opp Leader Service Line', 'Contact'
    ]

    for col in cols_to_initialize:
        df[col] = None

    return dataframe_ordering(df)


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

        names_dict = {}

        # Step 1: Find names in df['Opportunity Originator'] that are not in originator_df['Full Name']
        missing_names = list(df[~df['Opportunity Originator'].isin(originator_df['Full Name'])]['Opportunity Originator'])

        for name in missing_names:
            names_dict[name] = {'type': 'missing', 'Department (Advisory)': None}

        # Step 2: Find names in df['Opportunity Originator'] that are in originator_df['Full Name']
        # and have a null value in originator_df['Department (Advisory Report)']
        merged_df = df.merge(originator_df, left_on='Opportunity Originator', right_on='Full Name', how='inner')
        found_names_null_department = list(merged_df[merged_df['Department (Advisory Report)'].isnull()]['Opportunity Originator'])

        for name in found_names_null_department:
            names_dict[name] = {'type': 'null', 'Department (Advisory)': None}

        if len(names_dict):
            # Prompt user to determine the 'ADV?' value for each unique unmatched originator
            new_entries = prompt_adv_values(names_dict)

            for name in new_entries:
                if new_entries[name]['type'] == 'missing':
                    new_row = pd.DataFrame([{
                        'Full Name': name,
                        'Company': None,
                        'Department': None,
                        'Department (Advisory Report)': new_entries[name]['Department (Advisory)'],
                        'Department (Outsourced Report)': None,
                        'Job Title': None,
                        'Office Location': None,
                        'Date Updated': datetime.now().strftime("%Y-%m-%d")
                    }])
                    # Adding a new row using pd.concat
                    originator_df = pd.concat([originator_df, new_row], ignore_index=True)
                else:
                    # Update the 'Department (Outsourced Report)' for 'John Smith'
                    originator_df.loc[
                        originator_df['Full Name'] == name, ['Department (Advisory Report)', 'Date Updated']
                    ] = [new_entries[name]['Department (Advisory)'], datetime.now().strftime("%Y-%m-%d")]

            originator_df = originator_df.sort_values(by='Full Name').reset_index(drop=True)

        try:
            originator_df.to_excel(
                f'Reporting Output Files (Updated {datetime.now().strftime("%Y-%m-%d")})/Originators List.xlsx',
                index=False
            )
        except PermissionError:
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror(
                "Permission Error",
                "The file 'Originators List.xlsx' appears to be open. Please close the file and press OK."
            )
            root.destroy()
            originator_df.to_excel(
                f'Reporting Output Files (Updated {datetime.now().strftime("%Y-%m-%d")})/Originators List.xlsx',
                index=False
            )

        # Merge the two dataframes on the 'Opportunity Originator' column
        df = pd.merge(
            df,
            originator_df[['Full Name', 'Department (Advisory Report)']],
            left_on='Opportunity Originator',
            right_on='Full Name',
            how='left'
        )

        df.drop(columns=['Full Name'], inplace=True)

        df.rename(columns={'Department (Advisory Report)': 'ADV?'}, inplace=True)

        df.loc[:, 'Stage (adjusted)'] = pd.Categorical(
            df['Stage (adjusted)'],
            categories=['Closed Won', 'Closed Lost'],
            ordered=True
        )

        df = df.sort_values(by=['Stage (adjusted)', 'Close Date'], ascending=[False, False])

        return df

    df.loc[:, 'Stage (adjusted)'] = pd.Categorical(
        df['Stage (adjusted)'],
        categories=['Proposal', 'Qualified', 'Unqualified', 'Suspect'],
        ordered=True
    )

    # Sort the dataframe
    df = df.sort_values(by=['Stage (adjusted)', 'Created Date'], ascending=[True, False])

    return df


def clean_combined_out(df, active=True):
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

    df['Opportunity ID'] = df['Opportunity ID'].astype(str)

    df['Created Date'] = pd.to_datetime(df['Created Date'], errors='coerce')
    df['Created Date'] = df['Created Date'].dt.strftime('%Y-%m-%d')

    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    df['Close Date'] = df['Close Date'].dt.strftime('%Y-%m-%d')

    df['Name (adjusted for pivot)'] = df['Account Name'] + ' - ' + df['Opportunity ID']

    return df
