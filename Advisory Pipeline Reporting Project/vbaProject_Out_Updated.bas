Attribute VB_Name = "Module1"
Sub RunAll()
    ReformatOUTActive
    ReformatOUTClosed
    CreateSummarySheet
    CreateActiveByServiceAndLeader
    CreateMMCASummary
End Sub


Sub ReformatOUTActive()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OUT Active")
    
    ' Hide columns A-D and U-AE
    ws.Columns("A:D").EntireColumn.Hidden = True
    ws.Columns("U:AE").EntireColumn.Hidden = True
    
    ' Set the height of the first row and enable text wrapping
    ws.Rows(1).RowHeight = 31
    ws.Rows(1).WrapText = True
    
    ' Enable filters for all columns
    ws.Cells.AutoFilter

    ' Set column widths
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 47
    ws.Columns("G").ColumnWidth = 47
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("J").ColumnWidth = 17
    ws.Columns("K").ColumnWidth = 15
    ws.Columns("L").ColumnWidth = 9
    ws.Columns("M").ColumnWidth = 18
    ws.Columns("N").ColumnWidth = 16
    ws.Columns("O").ColumnWidth = 29
    ws.Columns("P").ColumnWidth = 33
    ws.Columns("Q").ColumnWidth = 18
    ws.Columns("R").ColumnWidth = 32
    ws.Columns("S").ColumnWidth = 22
    ws.Columns("T").ColumnWidth = 16

    ' Set currency format for columns H and I
    ws.Columns("H:I").NumberFormat = "$#,##0.00"
    
    ' Set short date format for columns J and K
    ws.Columns("J:K").NumberFormat = "mm/dd/yyyy"

End Sub

Sub ReformatOUTClosed()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("OUT Closed")
    
    ' Hide columns A-D and U-AE
    ws.Columns("A:D").EntireColumn.Hidden = True
    ws.Columns("U:AE").EntireColumn.Hidden = True
    
    ' Set the height of the first row and enable text wrapping
    ws.Rows(1).RowHeight = 31
    ws.Rows(1).WrapText = True
    
    ' Enable filters for all columns
    ws.Cells.AutoFilter

    ' Set column widths
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 47
    ws.Columns("G").ColumnWidth = 47
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 18
    ws.Columns("J").ColumnWidth = 17
    ws.Columns("K").ColumnWidth = 15
    ws.Columns("L").ColumnWidth = 9
    ws.Columns("M").ColumnWidth = 18
    ws.Columns("N").ColumnWidth = 16
    ws.Columns("O").ColumnWidth = 29
    ws.Columns("P").ColumnWidth = 33
    ws.Columns("Q").ColumnWidth = 18
    ws.Columns("R").ColumnWidth = 32
    ws.Columns("S").ColumnWidth = 22
    ws.Columns("T").ColumnWidth = 16

    ' Set currency format for columns H and I
    ws.Columns("H:I").NumberFormat = "$#,##0.00"
    
    ' Set short date format for columns J and K
    ws.Columns("J:K").NumberFormat = "mm/dd/yyyy"

End Sub

Sub CreateSummarySheet()
    Dim wb As Workbook
    Dim wsSummary As Worksheet
    Dim wsOutActive As Worksheet
    Dim wsOutClosed As Worksheet
    Dim pt As PivotTable
    Dim ptCache As PivotCache
    Dim dataRange As Range

    ' Set workbook and worksheet variables
    Set wb = ThisWorkbook
    Set wsOutActive = wb.Sheets("OUT Active")
    Set wsOutClosed = wb.Sheets("OUT Closed")
    
    ' Create a new worksheet for Summary
    On Error Resume Next
    Set wsSummary = wb.Sheets("Summary (All OUT)")
    On Error GoTo 0
    
    ' Check if the Summary sheet already exists and delete if necessary
    If Not wsSummary Is Nothing Then
        Application.DisplayAlerts = False
        wsSummary.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Add the Summary sheet as the first sheet
    Set wsSummary = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsSummary.Name = "Summary (All OUT)"
    
    ' Put the text 'Active Opportunities (All Outsourced)' in cell A2, bold
    wsSummary.Range("A2").Value = "Active Opportunities (All Outsourced)"
    wsSummary.Range("A2").Font.Bold = True
    
    ' Set data range for the first pivot table
    Set dataRange = wsOutActive.Range("E1:T" & wsOutActive.Cells(wsOutActive.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the first pivot cache and table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsSummary.Range("A3"), TableName:="SummaryPivotTable")
    
    ' Set up the first pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Type").Orientation = xlRowField
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Refresh the pivot table
        .RefreshTable
    End With
    
    ' Filter out "Renewal Business" in the row labels
    With wsSummary.PivotTables("SummaryPivotTable").PivotFields("Type")
        .PivotItems("Renewal Business").Visible = False
    End With

    ' Add the text 'FY Wins/Losses' to A27, bold
    wsSummary.Range("A27").Value = "FY Wins/Losses"
    wsSummary.Range("A27").Font.Bold = True
    
    ' Set data range for the second pivot table
    Set dataRange = wsOutClosed.Range("E1:T" & wsOutClosed.Cells(wsOutClosed.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the second pivot cache and table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsSummary.Range("A28"), TableName:="WinsLossesPivotTable")
    
    ' Set up the second pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        .PivotFields("Type").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Refresh the pivot table
        .RefreshTable
    End With

End Sub

Sub CreateActiveByServiceAndLeader()
    Dim wb As Workbook
    Dim wsActiveBySvc As Worksheet
    Dim wsOutActive As Worksheet
    Dim pt As PivotTable
    Dim ptCache As PivotCache
    Dim dataRange As Range

    ' Set workbook and worksheet variables
    Set wb = ThisWorkbook
    Set wsOutActive = wb.Sheets("OUT Active")
    
    ' Create a new worksheet for Active, By Svc & Leader
    On Error Resume Next
    Set wsActiveBySvc = wb.Sheets("Active, By Svc & Leader")
    On Error GoTo 0
    
    ' Check if the sheet already exists and delete if necessary
    If Not wsActiveBySvc Is Nothing Then
        Application.DisplayAlerts = False
        wsActiveBySvc.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Add the sheet as the last sheet
    Set wsActiveBySvc = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsActiveBySvc.Name = "Active, By Svc & Leader"
    
    ' Put the text 'Active Opportunities (All Outsourced)' in cell A2, bold
    wsActiveBySvc.Range("A2").Value = "Active Opportunities (All Outsourced)"
    wsActiveBySvc.Range("A2").Font.Bold = True
    
    ' Set data range for the pivot table
    Set dataRange = wsOutActive.Range("E1:AE" & wsOutActive.Cells(wsOutActive.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the pivot cache and table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsActiveBySvc.Range("A3"), TableName:="ActiveBySvcLeaderPivotTable")
    
    ' Set up the pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Service Lines").Orientation = xlRowField
        .PivotFields("Opportunity Leader").Orientation = xlRowField
        .PivotFields("Name (adjusted for pivot)").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Refresh the pivot table to ensure it is fully updated
        .RefreshTable
        
        ActiveSheet.PivotTables("ActiveBySvcLeaderPivotTable").PivotFields( _
        "Service Lines").ShowDetail = False
        
    End With

End Sub

Sub CreateMMCASummary()
    Dim wb As Workbook
    Dim wsMMCAS As Worksheet
    Dim wsOutActive As Worksheet
    Dim wsOutClosed As Worksheet
    Dim pt As PivotTable
    Dim ptCache As PivotCache
    Dim dataRange As Range

    ' Set workbook and worksheet variables
    Set wb = ThisWorkbook
    Set wsOutActive = wb.Sheets("OUT Active")
    Set wsOutClosed = wb.Sheets("OUT Closed")
    
    ' Create a new worksheet for MM CAS Summary
    On Error Resume Next
    Set wsMMCAS = wb.Sheets("MM CAS Summary")
    On Error GoTo 0
    
    ' Check if the sheet already exists and delete if necessary
    If Not wsMMCAS Is Nothing Then
        Application.DisplayAlerts = False
        wsMMCAS.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Add the new sheet at the end of the workbook
    Set wsMMCAS = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsMMCAS.Name = "MM CAS Summary"
    
    ' Put the text 'MidMarket CAS - Active Opportunities' in cell A1, bold
    wsMMCAS.Range("A1").Value = "MidMarket CAS - Active Opportunities"
    wsMMCAS.Range("A1").Font.Bold = True
    
    ' Set data range for the first pivot table
    Set dataRange = wsOutActive.Range("E1:AE" & wsOutActive.Cells(wsOutActive.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the pivot cache and table for the first pivot table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsMMCAS.Range("A4"), TableName:="MMCASPivotTable")
    
    ' Set up the first pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Type").Orientation = xlRowField
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Add filter for Service Lines
        With .PivotFields("Service Lines")
            .Orientation = xlPageField
            .Position = 1
        End With
        
        ' Refresh the pivot table
        .RefreshTable
    End With
    
    ' Filter for Service Lines
    With wsMMCAS.PivotTables("MMCASPivotTable").PivotFields("Service Lines")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        Dim serviceLine As Variant
        For Each serviceLine In Array("DFIR/Security Consulting", "IT Infrastructure", "IT Support", _
                                      "M - Mgmt - Accounting Consulting", "Maintenance Renewal", "Managed Services", _
                                      "Managed Tech", "OIT (Unknown)", "OUT Advanced Accounting", _
                                      "OUT Family Office", "OUT Fund Admin", "OUT IT Cloud Application Services", _
                                      "OUT IT Cyber Security Recurring Services", "OUT IT Cyber Security Services", _
                                      "OUT IT Recurring Services", "OUT IT Services", "OUT Outsourced Accounting Services", _
                                      "OUT Outsourced Solut for Financial Svcs", "OUT Property Accounting", "OUT RESIG", _
                                      "OUT Start Up – Accounting", "OUT Start Up - Finance", "OUT Start Up - HR", _
                                      "PAMI - Association Management", "Product", "Startup", "TECH - CYBERVEIL", _
                                      "TECH - IT Infrastructure; TECH - IT Support; TECH - Managed Services", "TECH - IT Support")
            On Error Resume Next
            .PivotItems(serviceLine).Visible = False
            On Error GoTo 0
        Next serviceLine
    End With

    ' Filter out Renewal Business in Type
    With wsMMCAS.PivotTables("MMCASPivotTable").PivotFields("Type")
        On Error Resume Next
        .PivotItems("Renewal Business").Visible = False
        On Error GoTo 0
    End With

    ' Add the text 'MidMarket CAS - FY Wins/Losses' in cell A20, bold
    wsMMCAS.Range("A20").Value = "MidMarket CAS - FY Wins/Losses"
    wsMMCAS.Range("A20").Font.Bold = True
    
    ' Set data range for the second pivot table
    Set dataRange = wsOutClosed.Range("E1:AE" & wsOutClosed.Cells(wsOutClosed.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the pivot cache and table for the second pivot table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsMMCAS.Range("A24"), TableName:="MMCASWinsLossesPivotTable")
    
    ' Set up the second pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        .PivotFields("Type").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Add filter for Service Lines
        With .PivotFields("Service Lines")
            .Orientation = xlPageField
            .Position = 1
        End With
        
        ' Refresh the pivot table
        .RefreshTable
    End With
    
    ' Filter for Service Lines in the second pivot table
    With wsMMCAS.PivotTables("MMCASWinsLossesPivotTable").PivotFields("Service Lines")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each serviceLine In Array("DFIR/Security Consulting", "IT Infrastructure", "IT Support", _
                                      "M - Mgmt - Accounting Consulting", "Maintenance Renewal", "Managed Services", _
                                      "Managed Tech", "OIT (Unknown)", "OUT Advanced Accounting", _
                                      "OUT Family Office", "OUT Fund Admin", "OUT IT Cloud Application Services", _
                                      "OUT IT Cyber Security Recurring Services", "OUT IT Cyber Security Services", _
                                      "OUT IT Recurring Services", "OUT IT Services", "OUT Outsourced Accounting Services", _
                                      "OUT Outsourced Solut for Financial Svcs", "OUT Property Accounting", "OUT RESIG", _
                                      "OUT Start Up – Accounting", "OUT Start Up - Finance", "OUT Start Up - HR", _
                                      "PAMI - Association Management", "Product", "Startup", "TECH - CYBERVEIL", _
                                      "TECH - IT Infrastructure; TECH - IT Support; TECH - Managed Services", "TECH - IT Support")
            On Error Resume Next
            .PivotItems(serviceLine).Visible = False
            On Error GoTo 0
        Next serviceLine
    End With

End Sub

Sub CreateMMCASOriginations()
    Dim wb As Workbook
    Dim wsMMCASOriginations As Worksheet
    Dim wsOutActive As Worksheet
    Dim pt As PivotTable
    Dim ptCache As PivotCache
    Dim dataRange As Range

    ' Set workbook and worksheet variables
    Set wb = ThisWorkbook
    Set wsOutActive = wb.Sheets("OUT Active")
    
    ' Create a new worksheet for MM CAS Originations
    On Error Resume Next
    Set wsMMCASOriginations = wb.Sheets("MM CAS Originations")
    On Error GoTo 0
    
    ' Check if the sheet already exists and delete if necessary
    If Not wsMMCASOriginations Is Nothing Then
        Application.DisplayAlerts = False
        wsMMCASOriginations.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Add the new sheet at the end of the workbook
    Set wsMMCASOriginations = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsMMCASOriginations.Name = "MM CAS Originations"
    
    ' Set data range for the pivot table
    Set dataRange = wsOutActive.Range("E1:AE" & wsOutActive.Cells(wsOutActive.Rows.Count, "E").End(xlUp).Row)
    
    ' Create the pivot cache and table
    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsMMCASOriginations.Range("A5"), TableName:="MMCASOriginationsPivotTable")
    
    ' Set up the pivot table fields
    With pt
        ' Add row fields
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Opportunity Originator").Orientation = xlRowField
        
        ' Add values
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
        
        ' Format the sum of First Year Fees as currency
        .PivotFields("Sum of First Year Fees").NumberFormat = "$#,##0.00"
        
        ' Add filters
        With .PivotFields("Service Lines")
            .Orientation = xlPageField
            .Position = 1
        End With
        With .PivotFields("Stage (adjusted)")
            .Orientation = xlPageField
            .Position = 2
        End With
        With .PivotFields("Type")
            .Orientation = xlPageField
            .Position = 3
        End With
        
        ' Refresh the pivot table
        .RefreshTable
    End With
    
    ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields( _
        "Service Lines").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields( _
        "Service Lines")
        .PivotItems("DFIR/Security Consulting").Visible = False
        .PivotItems("IT Infrastructure").Visible = False
        .PivotItems("IT Support").Visible = False
        .PivotItems("M - Mgmt - Accounting Consulting").Visible = False
        .PivotItems("Maintenance Renewal").Visible = False
        .PivotItems("Managed Services").Visible = False
        .PivotItems("Managed Tech").Visible = False
        .PivotItems("OIT (Unknown)").Visible = False
        .PivotItems("OUT Advanced Accounting").Visible = False
        .PivotItems("OUT Family Office").Visible = False
        .PivotItems("OUT Fund Admin").Visible = False
        .PivotItems("OUT IT Cloud Application Services").Visible = False
        .PivotItems("OUT IT Cyber Security Recurring Services").Visible = False
        .PivotItems("OUT IT Cyber Security Services").Visible = False
        .PivotItems("OUT IT Recurring Services").Visible = False
        .PivotItems("OUT IT Services").Visible = False
        .PivotItems("OUT Outsourced Accounting Services").Visible = False
        .PivotItems("OUT Outsourced Solut for Financial Svcs").Visible = False
        .PivotItems("OUT Property Accounting").Visible = False
        .PivotItems("OUT RESIG").Visible = False
    End With
    With ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields( _
        "Service Lines")
        .PivotItems("OUT Start Up – Accounting").Visible = False
        .PivotItems("OUT Start Up - Finance").Visible = False
        .PivotItems("OUT Start Up - HR").Visible = False
        .PivotItems("PAMI - Association Management").Visible = False
        .PivotItems("Product").Visible = False
        .PivotItems("Startup").Visible = False
        .PivotItems("TECH - CYBERVEIL").Visible = False
        .PivotItems( _
        "TECH - IT Infrastructure; TECH - IT Support; TECH - Managed Services"). _
        Visible = False
        .PivotItems("TECH - IT Support").Visible = False
    End With
    ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields("Type"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields("Type")
        .PivotItems("Renewal Business").Visible = False
    End With
    ActiveSheet.PivotTables("MMCASOriginationsPivotTable").PivotFields("Type"). _
        EnableMultiplePageItems = True

End Sub







