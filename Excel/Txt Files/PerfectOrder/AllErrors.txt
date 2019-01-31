Sub ALLERRORS()
'
' ALL ERRORS Macro
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
    ActiveWindow.SmallScroll Down:=3
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
    Range("A4").Select
    ActiveSheet.PivotTables("ALLERRORS").Name = "ALLERRORS"
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
------------------------------------------------------------------------------------
Sub ALLERRORS()						**9/27/2017**
'
' ALLERRORS Macro
' Creates the ALLERRORS Table
'
    Dim Sh As Worksheets
    Set AESh = Worksheets("All Errors")

    AESh.Activate
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        Position = 2
        PivotItems("").Visible = False
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        Position = 1
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        Position = 4
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        Position = 3
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible")
        Orientation = xlRowField
        Position = 5
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    ActiveSheet.PivotTables("ALLERRORS").AddDataField ActiveSheet.PivotTables( _
        "ALLERRORS").PivotFields("L1 Error"), "Count of L1 Error", xlCount
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        Position = 1
    End With
    With ActiveSheet.PivotTables("ALLERRORS")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A2").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub
----------------------------------------------------------------------------------
Sub ALLERRORS()            '**9/29/2017** - Edit subtotals
'                           
' ALLERRORS Macro
' Creates the ALLERRORS Table
'
    Dim Sh As Worksheets
    Set AESh = Worksheets("All Errors")

    AESh.Activate
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("").Visible = False
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible")
        Orientation = xlRowField
        .Position = 5
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
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
    Range("A2").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub