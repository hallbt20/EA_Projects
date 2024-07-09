import pandas as pd
import pyodbc

scheduled_df = pd.read_csv('Runn Month by Month for RCS.csv')

scheduled_df = pd.melt(
    scheduled_df,
    id_vars=['Full Name', 'Email', 'Default Role', 'Team', 'Person Status'],
    value_vars=[
        'Jan 2024 Scheduled Hours', 'Feb 2024 Scheduled Hours',
        'Mar 2024 Scheduled Hours', 'Apr 2024 Scheduled Hours',
        'May 2024 Scheduled Hours', 'Jun 2024 Scheduled Hours'],
    var_name='Month',
    value_name='Scheduled Hours'
)

month_replacements = {
    'Jan 2024 Scheduled Hours': 'January',
    'Feb 2024 Scheduled Hours': 'February',
    'Mar 2024 Scheduled Hours': 'March',
    'Apr 2024 Scheduled Hours': 'April',
    'May 2024 Scheduled Hours': 'May',
    'Jun 2024 Scheduled Hours': 'June'
}

scheduled_df['Month'] = scheduled_df['Month'].replace(month_replacements)

scheduled_df = scheduled_df[~scheduled_df['Email'].isna()]

# Define custom order for the 'Month' column
month_order = ['January', 'February', 'March', 'April', 'May', 'June']
scheduled_df['Month'] = pd.Categorical(scheduled_df['Month'], categories=month_order, ordered=True)

# Sort by 'Full Name' and then by the custom order of 'Month'
scheduled_df = scheduled_df.sort_values(by=['Full Name', 'Month'])

# Below is the code for gathering billable hours

# Define the connection string
conn_str = (
    r"Driver={ODBC Driver 17 for SQL Server};"
    r"Server=pncservice.pncpa.com\Reporting;"
    r"Database=AdvisoryDM;"
    r"Trusted_Connection=yes;"
)

# Establish the connection
conn = pyodbc.connect(conn_str)

# Define the SQL query
query = """
    WITH CTE AS (
      SELECT 
        [colleagueFullName],
        [colleagueEmail],
        [colleagueManagementLevel],
        [colleagueLocation],
        [colleaguePracticeArea],
        [colleaguePracticeGroup],
        [colleaguePractice],
        [colleagueStatus],
        [monthName],
        SUM([wipHours]) AS 'Billable Hours',
        ROW_NUMBER() OVER (PARTITION BY [colleagueFullName], [monthName] ORDER BY CASE WHEN [colleagueStatus] = 'Inactive' THEN 0 ELSE 1 END) AS rn
      FROM [AdvisoryDM].[dbo].[AllAdvisoryData] AAD
      JOIN [AdvisoryDM].[dbo].[dimColleague] DC ON employeeID = [colleagueLegacyEmployeeID]
      WHERE colleaguePracticeGroup = 'Risk & Compliance Services (RCS)'
        AND month BETWEEN 1 AND 6
        AND year = 2024
        AND wipIsBillable = 1
        AND colleaguePractice IN (
          'Internal Audit & GRC', 
          'IT Risk, Data Privacy & Security', 
          'Financial & Regulatory Risk Services',
          'ATCS'
        )
      GROUP BY 
        [colleagueFullName],
        [colleaguePractice],
        [colleagueEmail],
        [monthName],
        [colleagueManagementLevel],
        [colleagueLocation],
        [colleaguePracticeArea],
        [colleaguePracticeGroup],
        [colleaguePractice],
        [colleagueStatus]
    )
    SELECT 
      [colleagueFullName],
      [colleagueEmail],
      [colleagueManagementLevel],
      [colleagueLocation],
      [colleaguePracticeArea],
      [colleaguePracticeGroup],
      [colleaguePractice],
      [colleagueStatus],
      [monthName],
      [Billable Hours]
    FROM CTE
    WHERE rn = 1
    ORDER BY 
      [colleaguePractice],
      [colleagueFullName], 
      CASE 
        WHEN [monthName] = 'January' THEN 1
        WHEN [monthName] = 'February' THEN 2
        WHEN [monthName] = 'March' THEN 3
        WHEN [monthName] = 'April' THEN 4
        WHEN [monthName] = 'May' THEN 5
        WHEN [monthName] = 'June' THEN 6
      END;
"""

# Execute the query and fetch the data into a pandas DataFrame
billable_df = pd.read_sql_query(query, conn)

# Close the connection
conn.close()

df_merged = pd.merge(
    billable_df,
    scheduled_df,
    left_on=['colleagueEmail', 'monthName'],
    right_on=['Email', 'Month'],
    how='outer'
)

# Define custom order for the 'Month' column
month_order = ['January', 'February', 'March', 'April', 'May', 'June']
df_merged['Month'] = pd.Categorical(df_merged['Month'], categories=month_order, ordered=True)

# Sort by 'Full Name' and then by the custom order of 'Month'
df_merged = df_merged.sort_values(by=['colleagueFullName', 'Month'])

df_merged.to_excel('RCS - Scheduled VS Actual WIP Hours.xlsx', index=False)
