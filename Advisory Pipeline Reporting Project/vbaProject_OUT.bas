Attribute VB_Name = "Module1"
Sub FormatSheets()
    Dim wsActive As Worksheet
    Dim wsClosed As Worksheet
    Dim ws As Worksheet
    
    ' Set the source worksheets
    Set wsActive = ThisWorkbook.Sheets("OUT Active")
    Set wsClosed = ThisWorkbook.Sheets("OUT Closed")
    
    ' Format the "OUT Active" sheet
    Call FormatSheet(wsActive)
    
    ' Format the "OUT Closed" sheet
    Call FormatSheet(wsClosed)
End Sub

Sub FormatSheet(ws As Worksheet)
    With ws
        ' Turn on filters for all columns
        .AutoFilterMode = False
        .Range("A1").AutoFilter
        
        ' Extend the column widths for all columns
        .Cells.EntireColumn.AutoFit
        
        ' Hide columns A through D
        .Columns("A:D").Hidden = True
        .Columns("U:AE").Hidden = True
        
        ' Make columns H and I monetary types
        .Columns("H:I").NumberFormat = "$#,##0.00"
        
        ' Make columns J and K short date types
        .Columns("J:K").NumberFormat = "mm/dd/yyyy"
    End With
End Sub


Sub CreatePivotTables()
    Dim wsSummary As Worksheet
    Dim wsActive As Worksheet
    Dim wsClosed As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim wsSummaryExists As Boolean
    Dim ws As Worksheet
    
    ' Check if "Summary (All OUT)" sheet exists
    wsSummaryExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Summary (All OUT)" Then
            wsSummaryExists = True
            Exit For
        End If
    Next ws
    
    ' If the sheet exists, delete it and create a new one
    If wsSummaryExists Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets("Summary (All OUT)").Delete
        Application.DisplayAlerts = True
    End If
    
    ' Create the "Summary (All OUT)" sheet
    Set wsSummary = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsSummary.Name = "Summary (All OUT)"
    
    ' Set the source worksheets
    Set wsActive = ThisWorkbook.Sheets("OUT Active")
    Set wsClosed = ThisWorkbook.Sheets("OUT Closed")
    
    ' Add title and format it
    With wsSummary
        .Range("A2").Value = "Active Opportunities (All Outsourced)"
        .Range("A2").Font.Bold = True
        .Range("A25").Value = "FY Wins/Losses"
        .Range("A25").Font.Bold = True
    End With
    
    ' Create the first pivot table
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsActive.Range("A1:AE" & wsActive.Cells(wsActive.Rows.Count, "A").End(xlUp).Row))
    Set pt = pc.CreatePivotTable(TableDestination:=wsSummary.Range("A3"), TableName:="PivotTable1")
    
    With pt
        .PivotFields("Type").Orientation = xlRowField
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
    End With
    
    ' Create the second pivot table
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsClosed.Range("A1:AE" & wsClosed.Cells(wsClosed.Rows.Count, "A").End(xlUp).Row))
    Set pt = pc.CreatePivotTable(TableDestination:=wsSummary.Range("A26"), TableName:="PivotTable2")
    
    With pt
        .PivotFields("Stage (adjusted)").Orientation = xlRowField
        .PivotFields("Type").Orientation = xlRowField
        .AddDataField .PivotFields("Opportunity Name"), "Count of Opportunity Name", xlCount
        .AddDataField .PivotFields("First Year Fees"), "Sum of First Year Fees", xlSum
    End With
End Sub

