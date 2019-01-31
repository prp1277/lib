Sub ExecutionErrors()
'
' ExecutionErrorPivot Macro
' Creates the Execution Error Pivot Table
'

'
    Sheets.Add
    ActiveWorkbook.Worksheets("ALLERRORS").PivotTables("PlatinumPivot"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet5!R3C1", TableName:= _
        "PivotTable2", DefaultVersion:=xlPivotTableVersion15
    Sheets("Sheet5").Select
    Cells(3, 1).Select
    Sheets("Sheet5").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .PivotItems("Availability Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    Sheets("Sheet5").Name = "Execution Error Pivot"
    ActiveSheet.PivotTables("PivotTable2").Name = "ExecutionError"
    With ActiveSheet.PivotTables("ExecutionError")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Customer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Invoice  #").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("ExecutionError").PivotFields("Responsible").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub