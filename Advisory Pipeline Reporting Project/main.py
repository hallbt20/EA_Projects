import sys
import warnings

from clean_reports import *
import popups


def adv_process():
    # ADV Report start
    with warnings.catch_warnings(action='ignore'):
        all_active = pd.concat([sf_active, ns_active, great_lakes_active], ignore_index=True)
        all_closed = pd.concat([sf_closed, ns_closed, great_lakes_closed], ignore_index=True)

    # Split off Advisory data
    adv_active = clean_combined_adv(all_active)
    adv_closed = clean_combined_adv(all_closed, report_files[5], False)

    originator_list_df = pd.read_excel(report_files[5])

    file_name = f'ADV Pipeline (Updated on {date_now})'

    try:
        os.chdir(sys._MEIPASS)
    except AttributeError:
        pass

    with pd.ExcelWriter(f'{cwd}/{file_name}.xlsx', engine='xlsxwriter') as writer:
        # Write each dataframe to a different sheet
        adv_active.to_excel(writer, sheet_name='ADV Active', index=False)
        adv_closed.to_excel(writer, sheet_name='ADV Closed', index=False)
        originator_list_df.to_excel(writer, sheet_name='Originators List', index=False)

        worksheet = writer.sheets['Originators List']
        worksheet.hide()

    with pd.ExcelWriter(f'{cwd}/temp.xlsx', engine='xlsxwriter') as writer:
        # Write each dataframe to a different sheet
        adv_active.to_excel(writer, sheet_name='ADV Active', index=False)
        adv_closed.to_excel(writer, sheet_name='ADV Closed', index=False)
        originator_list_df.to_excel(writer, sheet_name='Originators List', index=False)

        worksheet = writer.sheets['Originators List']
        worksheet.hide()

        workbook = writer.book
        workbook.filename = f'{cwd}/{file_name}.xlsm'
        workbook.add_vba_project('./vbaProject.bin')

    os.remove(f'{cwd}/temp.xlsx')

    # Call the function at the end of script
    show_report_generated_message(file_name, cwd)

    os.chdir(cwd)


def out_process():
    # Outsourcing for Salesforce
    sf_active_out = sf_active[sf_active['Service Line Group'] == 'OUT']
    sf_closed_out = sf_closed[sf_closed['Service Line Group'] == 'OUT']

    # Outsourcing for Netsuite
    ns_active_out = ns_active[ns_active['Service Line Group'] == 'OUT (BOutsourced Services)']
    ns_closed_out = ns_closed[ns_closed['Service Line Group'] == 'OUT (BOutsourced Services)']

    try:
        hubspot_df = pd.read_excel(report_files[6])
    except ValueError:
        hubspot_df = pd.read_csv(report_files[6])

    hubspot_df = clean_hubspot(hubspot_df)
    hubspot_active = hubspot_df[~(hubspot_df['Stage (adjusted)'] == 'Closed Lost')]
    # Filter the DataFrame
    hubspot_closed = hubspot_df[
        (hubspot_df['Stage (adjusted)'] == 'Closed Lost') &
        (hubspot_df['Close Date'] >= pd.to_datetime('2023-08-01'))
    ]

    # Read Gulf Coast Outsourced IT Data
    eag_gc_oit = pd.read_excel(report_files[7])

    pnt_cw_active = clean_pnt(eag_gc_oit[eag_gc_oit['Closed_Status'].isna()])
    pnt_cw_closed = clean_pnt(
        eag_gc_oit[
            (~eag_gc_oit['Closed_Status'].isna()) &
            (~eag_gc_oit['Closed_Date'].isna()) &
            (eag_gc_oit['Closed_Date'] >= pd.Timestamp('2023-08-01')) &
            (eag_gc_oit['Closed_Date'] <= pd.Timestamp('now'))
            ], False
    )

    legacy_oit = clean_legacy(pd.read_excel(report_files[8]))

    triangle_file = report_files[9]
    triangle_active_sheet = find_sheet_name(triangle_file, 'Triangle Active')
    triangle_closed_sheet = find_sheet_name(triangle_file, 'Triangle Wins')

    triangle_active = clean_triangle(pd.read_excel(triangle_file, sheet_name=triangle_active_sheet))
    triangle_closed = clean_triangle(pd.read_excel(triangle_file, sheet_name=triangle_closed_sheet))

    # ADV Report start
    with warnings.catch_warnings(action='ignore'):
        out_active = pd.concat([
            sf_active_out,
            ns_active_out,
            hubspot_active,
            pnt_cw_active,
            legacy_oit,
            triangle_active
        ], ignore_index=True)

        out_closed = pd.concat([
            sf_closed_out,
            ns_closed_out,
            hubspot_closed,
            pnt_cw_closed,
            triangle_closed
        ])

        out_active = clean_combined_out(out_active)
        out_closed = clean_combined_out(out_closed)

    with pd.ExcelWriter(f'{cwd}/Out Pipeline.xlsx', engine='xlsxwriter') as writer:
        # Write each dataframe to a different sheet
        out_active.to_excel(writer, sheet_name='OUT Active', index=False)
        out_closed.to_excel(writer, sheet_name='OUT Closed', index=False)


if __name__ == "__main__":
    # Starts the UI window for selecting type of report and folder with corresponding files
    report_files = popups.startup_window()

    # Read, cleans, and reformats Salesforce data to a Pandas dataframe
    sf_active = clean_salesforce(pd.read_excel(report_files[0]))
    sf_closed = clean_salesforce(pd.read_excel(report_files[1]))

    # Read, clean, and NetSuite data as Pandas dataframes and returns cleaned data frame
    ns_active = clean_netsuite(pd.read_excel(report_files[2]))
    ns_closed = clean_netsuite(pd.read_excel(report_files[3]), False)

    great_lakes = clean_great_lakes(pd.read_excel(report_files[4]))
    great_lakes_active = great_lakes[great_lakes['Stage (adjusted)'].isin(['Proposal', 'Qualified', 'Unqualified'])]
    great_lakes_closed = great_lakes[great_lakes['Stage (adjusted)'].isin(['Closed Won', 'Closed Lost'])]

    # Create folder to store reporting files in
    date_now = datetime.now().strftime("%Y-%m-%d")
    new_directory = f'Reporting Output Files (Updated {date_now})'

    try:
        os.mkdir(new_directory)
    except FileExistsError:
        pass

    cwd = f'{os.getcwd()}/{new_directory}'

    if 'Advisory' in popups.selected_reports:
        adv_process()

    if 'Outsourcing' in popups.selected_reports:
        out_process()
