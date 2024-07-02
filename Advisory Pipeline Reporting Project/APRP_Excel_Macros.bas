Attribute VB_Name = "Module1"
Sub CreatePivotTables()
    Dim wsActive As Worksheet
    Dim wsClosed As Worksheet
    Dim wsSummary As Worksheet
    Dim ptCache1 As PivotCache
    Dim ptCache2 As PivotCache
    Dim pt1 As PivotTable
    Dim pt2 As PivotTable
    Dim dataRange1 As Range
    Dim dataRange2 As Range

    ' Set references to the relevant sheets
    Set wsActive = ThisWorkbook.Sheets("ADV Active")
    Set wsClosed = ThisWorkbook.Sheets("ADV Closed")

    ' Enable filtering on all columns in "ADV Active" and "ADV Closed"
    If Not wsActive.AutoFilterMode Then
        wsActive.Rows(1).AutoFilter
    End If
    If Not wsClosed.AutoFilterMode Then
        wsClosed.Rows(1).AutoFilter
    End If

    ' Auto-fit columns in "ADV Active" and "ADV Closed"
    wsActive.Cells.EntireColumn.AutoFit
    wsClosed.Cells.EntireColumn.AutoFit

    ' Create the Summary sheet if it doesn't exist
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSummary.Name = "Summary"
    Else
        ' Clear the existing sheet
        wsSummary.Cells.Clear
    End If
    
    ' Move the Summary sheet to the first position
    wsSummary.Move Before:=ThisWorkbook.Sheets(1)

    ' Define the data ranges
    Set dataRange1 = wsActive.Range("A1:T" & wsActive.Cells(wsActive.Rows.Count, "A").End(xlUp).Row)
    Set dataRange2 = wsClosed.Range("A1:U" & wsClosed.Cells(wsClosed.Rows.Count, "A").End(xlUp).Row)

    ' Create the PivotCache for the first PivotTable
    Set ptCache1 = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange1)

    ' Create the first PivotTable
    Set pt1 = ptCache1.CreatePivotTable( _
        TableDestination:=wsSummary.Range("A3"), _
        TableName:="PivotTableSummary1")
        
    ' Add title for the first PivotTable
    wsSummary.Range("A2").Value = "Active Opportunities"
    wsSummary.Range("A2").Font.Bold = True

    ' Set up the first PivotTable fields
    With pt1
        ' Rows
        With .PivotFields("Stage (adjusted)")
            .Orientation = xlRowField
            .Position = 1
        End With

        With .PivotFields("Type")
            .Orientation = xlRowField
            .Position = 2
        End With

        ' Values
        With .PivotFields("Opportunity Name")
            .Orientation = xlDataField
            .Function = xlCount
            .Position = 1
            .Caption = "Opportunity Count"
        End With

        With .PivotFields("First Year Fees")
            .Orientation = xlDataField
            .Function = xlSum
            .Position = 2
            .Caption = "First Year Fees "
            .NumberFormat = "$#,##0.00" ' Format as monetary with comma separators
        End With

        With .PivotFields("Age")
            .Orientation = xlDataField
            .Function = xlAverage
            .Position = 3
            .Caption = "Avg Age"
            .NumberFormat = "0" ' Round to the nearest whole number
        End With
    End With

    ' Add title for the second PivotTable
    wsSummary.Range("A26").Value = "FYTD Wins/Losses"
    wsSummary.Range("A26").Font.Bold = True

    ' Create the PivotCache for the second PivotTable
    Set ptCache2 = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange2)

    ' Create the second PivotTable
    Set pt2 = ptCache2.CreatePivotTable( _
        TableDestination:=wsSummary.Range("A27"), _
        TableName:="PivotTableSummary2")

    ' Set up the second PivotTable fields
    With pt2
        ' Rows
        With .PivotFields("Stage (adjusted)")
            .Orientation = xlRowField
            .Position = 1
        End With

        With .PivotFields("Type")
            .Orientation = xlRowField
            .Position = 2
        End With

        ' Values
        With .PivotFields("Opportunity Name")
            .Orientation = xlDataField
            .Function = xlCount
            .Position = 1
            .Caption = "Opportunity Count"
        End With

        With .PivotFields("First Year Fees")
            .Orientation = xlDataField
            .Function = xlSum
            .Position = 2
            .Caption = "First Year Fees "
            .NumberFormat = "$#,##0.00" ' Format as monetary with comma separators
        End With

        ' Add calculated field for % of Grand Total based on Opportunity Count
        With .PivotFields("Opportunity Name")
            .Orientation = xlDataField
            .Function = xlCount
            .Position = 3
            .Caption = "% of Grand Total - Opportunity Count"
            .Calculation = xlPercentOfTotal
        End With

        ' Add calculated field for % of Grand Total based on First Year Fees (EA's portion)
        With .PivotFields("First Year Fees")
            .Orientation = xlDataField
            .Function = xlSum
            .Position = 4
            .Caption = "% of Grand Total - First Year Fees"
            .Calculation = xlPercentOfTotal
        End With
    End With

    ' Autofit columns in the Summary sheet
    wsSummary.Columns.AutoFit

End Sub

Sub CreateAdditionalPivotTable()
    Dim wsClosed As Worksheet
    Dim wsOriginations As Worksheet
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

    ' Set reference to the relevant sheet
    Set wsClosed = ThisWorkbook.Sheets("ADV Closed")

    ' Create the Originations (Wins) sheet if it doesn't exist
    On Error Resume Next
    Set wsOriginations = ThisWorkbook.Sheets("Originations (Wins)")
    On Error GoTo 0
    If wsOriginations Is Nothing Then
        Set wsOriginations = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOriginations.Name = "Originations (Wins)"
    Else
        ' Clear the existing sheet
        wsOriginations.Cells.Clear
    End If
    
    ' Add title text in A1
    wsOriginations.Range("A1").Value = "Originations (Wins) by Individual (Grouped by ADV vs Other)"
    wsOriginations.Range("A1").Font.Bold = True

    ' Define the data range
    Set dataRange = wsClosed.Range("A1:U" & wsClosed.Cells(wsClosed.Rows.Count, "A").End(xlUp).Row)

    ' Create the PivotCache
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)

    ' Create the PivotTable
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=wsOriginations.Range("A4"), _
        TableName:="PivotTableOriginations")

    ' Set up the PivotTable fields
    With pt
        ' Filter
        With .PivotFields("Stage (adjusted)")
            .Orientation = xlPageField
            .Position = 1
        End With

        ' Rows
        With .PivotFields("Originator in ADV?")
            .Orientation = xlRowField
            .Position = 1
        End With

        With .PivotFields("Opportunity Originator")
            .Orientation = xlRowField
            .Position = 2
        End With

        ' Values
        With .PivotFields("Opportunity Name")
            .Orientation = xlDataField
            .Function = xlCount
            .Position = 1
            .Caption = "Opportunity Count"
        End With

        With .PivotFields("First Year Fees")
            .Orientation = xlDataField
            .Function = xlSum
            .Position = 2
            .Caption = "First Year Fees "
            .NumberFormat = "$#,##0.00" ' Format as monetary with comma separators
        End With
    End With

    ' Autofit columns in the Originations (Wins) sheet
    wsOriginations.Columns.AutoFit

End Sub

Sub HideColumns()
    Dim ws_active As Worksheet
    Dim ws_closed As Worksheet
    Set ws_active = ThisWorkbook.Sheets("ADV Active")
    Set ws_closed = ThisWorkbook.Sheets("ADV Closed")
    
    ' Hide columns A through D
    ws_active.Columns("A:D").Hidden = True
    ws_closed.Columns("A:D").Hidden = True
    
    ' Hide column N on ADV Closed
    ws_closed.Columns("N").Hidden = True
    
    ' Hide columns after Recurrence
    ws_active.Columns("U:AC").Hidden = True
    ws_closed.Columns("V:AD").Hidden = True
    
    ' Allow wrapping text in the first row
    Dim col As Range
    For Each col In ws_active.Range("A1:T1").Columns
        col.WrapText = True
        col.HorizontalAlignment = xlCenter
        col.VerticalAlignment = xlCenter
    Next col
    
    For Each col In ws_closed.Range("A1:U1").Columns
        col.WrapText = True
        col.HorizontalAlignment = xlCenter
        col.VerticalAlignment = xlCenter
    Next col
    
    ' Set the width of columns
    ws_active.Columns("E").ColumnWidth = 14
    ws_closed.Columns("E").ColumnWidth = 14
    ws_active.Columns("F").ColumnWidth = 36
    ws_closed.Columns("F").ColumnWidth = 36
    ws_active.Columns("G").ColumnWidth = 36
    ws_closed.Columns("G").ColumnWidth = 36
    
    ' Format columns H and I as currency
    ws_active.Columns("H").NumberFormat = "$#,##0.00"
    ws_closed.Columns("H").NumberFormat = "$#,##0.00"
    ws_active.Columns("I").NumberFormat = "$#,##0.00"
    ws_closed.Columns("I").NumberFormat = "$#,##0.00"
    
    ' Format column J to show only dates
    ws_active.Columns("J").NumberFormat = "mm/dd/yyyy"
    ws_closed.Columns("J").NumberFormat = "mm/dd/yyyy"
    ws_active.Columns("K").NumberFormat = "mm/dd/yyyy"
    ws_closed.Columns("K").NumberFormat = "mm/dd/yyyy"
    
    ' Set the width of columns
    ws_active.Columns("M").ColumnWidth = 27
    ws_closed.Columns("M").ColumnWidth = 27
    ws_active.Columns("N").ColumnWidth = 27
    ws_active.Columns("O").ColumnWidth = 35
    ws_closed.Columns("O").ColumnWidth = 35
    ws_active.Columns("P").ColumnWidth = 38
    ws_closed.Columns("Q").ColumnWidth = 38
    ws_active.Columns("R").ColumnWidth = 30
    ws_closed.Columns("S").ColumnWidth = 30
    ws_active.Columns("S").ColumnWidth = 20
    ws_closed.Columns("T").ColumnWidth = 20
    

End Sub
