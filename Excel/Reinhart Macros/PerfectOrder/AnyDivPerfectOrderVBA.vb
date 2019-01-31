Sub PivotData()
'
' PivotFormat Macro
' Formats the SQL Export into Pivot-Ready data
'

'
    Sheets("Invoices").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1)), TrailingMinusNumbers:=True
    Range("B5").Select
    Sheets.Add Before:=Worksheets(1)
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Invoices!R5C1:R1000C17", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="ALLERRORS", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets(1).Select
    Sheets(1).Name = "ALL ERRORS"
    Sheets("Macro").Select
End Sub
-------------------------------------------------------------------------------------
Sub ALLERRORS()
'
' ALLERRORS Macro
' Creates the ALLERRORS Table
'

'
    Sheets("ALL ERRORS").Select
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .PivotItems("").Visible = False
    End With
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("ALLERRORS").AddDataField ActiveSheet.PivotTables( _
        "ALLERRORS").PivotFields("L1 Error"), "Count of L1 Error", xlCount
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("ALLERRORS")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 3
    Sheets("Macro").Select
    End With
End Sub
---------------------------------------------------------------------------------------
Sub ExecutionErrors()
'
' ExecutionErrorPivot Macro
' Creates the Execution Error Pivot Table
'

'
    Sheets.Add Before:=Worksheets(1)
    Sheets("ALL ERRORS").Select
    ActiveWorkbook.Worksheets("ALL ERRORS").PivotTables("ALLERRORS"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet2!R3C1", TableName:= _
        "EXECUTIONERRORS", DefaultVersion:=xlPivotTableVersion15
    Sheets(1).Select
    Sheets(1).Name = "Execution Errors Pivot"
    Cells(3, 1).Select
    Sheets("Execution Errors Pivot").Select
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .PivotItems("Availability Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("EXECUTIONERRORS").AddDataField ActiveSheet.PivotTables( _
        "EXECUTIONERRORS").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ExecutionErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("ExecutionErrors").PivotFields("Customer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ExecutionErrors").PivotFields("Invoice Date"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("ExecutionErrors").PivotFields("Invoice  #").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("ExecutionErrors").PivotFields("Responsible").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    Cells.Select
    Cells.EntireColumn.AutoFit
    Sheets("Macro").Select
End Sub
-------------------------------------------------------------------------------------
Sub AvailabilityErrors()
'
' AvailabilityErrors Macro
' Creates the Availability Errors Tab and Pivot Table
'

'
    Sheets.Add Before:=Worksheets(1)
'Create and names the Availability Errors Pivot
    ActiveWorkbook.Worksheets("ALL ERRORS").PivotTables("ALLERRORS"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet3!R3C1", TableName:= _
        "AVAILABILITYERRORS", DefaultVersion:=xlPivotTableVersion15
'Using Sheets(1) adds it to the front of the workbook and makes it easier to reference
    Sheets(1).Select
    Cells(3, 1).Select
    Sheets(1).Select
'Name the new tab
    Sheets(1).Name = "Availability Errors"
'These could be combined with the .Subtotals below to run faster
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L1 Error")
        .PivotItems("Execution Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("AVAILABILITYERRORS").AddDataField ActiveSheet.PivotTables( _
        "AVAILABILITYERRORS").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("AVAILABILITYERRORS").Name = "AvailabilityErrors"
'Set the formatting to old school pivot table
    With ActiveSheet.PivotTables("AvailabilityErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
'Hide the subtotals of each field
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
'Autowidth the sheet
    Cells.EntireColumn.AutoFit
    Sheets("Macro").Select
End Sub
