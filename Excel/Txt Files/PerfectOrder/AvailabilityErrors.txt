Sub AvailabilityErrors()
'
' AvailabilityErrors Macro
' Creates the Availability Errors Tab and Pivot Table
'

'
    Sheets.Add
    ActiveWorkbook.Worksheets("ALLERRORS").PivotTables("PlatinumPivot"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet6!R3C1", TableName:= _
        "PivotTable3", DefaultVersion:=xlPivotTableVersion15
    Sheets("Sheet6").Select
    Cells(3, 1).Select
    Sheets("Sheet6").Select
    Sheets("Sheet6").Name = "Availability Errors"
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .PivotItems("Execution Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable3").Name = "AvailabilityErrors"
    With ActiveSheet.PivotTables("AvailabilityErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Customer"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("A#").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Invoice  #"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveSheet.PivotTables("AvailabilityErrors").PivotFields("Customer"). _
        Subtotals = Array(True, False, False, False, False, False, False, False, False, False, _
        False, False)
    Cells.EntireColumn.AutoFit
End Sub