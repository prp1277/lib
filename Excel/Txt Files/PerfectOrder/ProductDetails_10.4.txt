Sub ProductDetails()
'
' ProductDetails Macro
' Copy the All Errors Pivot and reformat it to focus on errors per product
'

'
    Sheets("ALL ERRORS").Select
    Sheets("ALL ERRORS").Copy After:=Sheets(6)
    Sheets("ALL ERRORS (2)").Select
    Sheets("ALL ERRORS (2)").Name = "Product Details"
    Range("I894").Select
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Description")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Description")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 7
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#").Orientation = xlHidden
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible").Orientation = _
        xlHidden
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L2 Error")
        .Orientation = xlRowField
        .Position = 5
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Count of L1 Error"). _
        Orientation = xlHidden
    ActiveSheet.PivotTables("ALLERRORS").AddDataField ActiveSheet.PivotTables( _
        "ALLERRORS").PivotFields("Product #"), "Count of Product #", xlCount
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("L3 Error").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("L2 Error").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Product #").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    Cells.Select
    Range("A855").Activate
    Cells.EntireColumn.AutoFit
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Product #").Subtotals = Array _
        (True, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Description").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Cells.EntireColumn.AutoFit
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Vendor Name")
        .Orientation = xlRowField
        .Position = 6
    End With
    ActiveSheet.PivotTables("ALLERRORS").PivotFields("Vendor Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Range("B855").Activate
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").Select
    Range("C855").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:G").Select
    Range("D855").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    Range("H855").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
End Sub
