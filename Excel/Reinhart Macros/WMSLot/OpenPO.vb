Sub OpenPO()
'
' OpenPO Macro
' Separate POs by buyer and format pivot table
'

'
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    
'Select all and form a table called OpenPO
    Range("A7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$7:$AB$500"), , xlYes).Name = _
        "OpenPO"
    Range("A7:AB351").Select
    Range("AC7").Select
    ActiveCell.FormulaR1C1 = "PO Comments"
    
'This is where I copied and pasted from One Note
    Range("AC8").Select
    ActiveSheet.Paste
    Range("AD7").Select
    ActiveCell.FormulaR1C1 = "Vendor Comment"
    Range("AD8").Select
    ActiveSheet.Paste
    Columns("AC:AD").EntireColumn.AutoFit
    
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=11, Criteria1:= _
        "B"
    Range("K8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "2"
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=11, Criteria1:= _
        "E"
    Range("K10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "1"
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=11, Criteria1:= _
        "N"
    Range("K122").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "0"
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=11
    
    
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=12, Criteria1:= _
        "Y"
    Range("L8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "2"
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=12, Criteria1:= _
        "N"
    Range("L122").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "1"
    ActiveWindow.SmallScroll Down:=-24
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=12
    
    
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=13, Criteria1:= _
        "Y"
    Range("M8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "2"
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=13, Criteria1:= _
        "N"
    Range("M54").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "1"
    Range("M7").Select
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=13
    ActiveWindow.SmallScroll Down:=-24
    
    
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=14, Criteria1:= _
        "Y"
    Range("N8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "2"
    Range("N7").Select
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=14, Criteria1:= _
        "N"
    Range("N30").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormulaR1C1 = "1"
    Range("N7").Select
    ActiveSheet.ListObjects("OpenPO").Range.AutoFilter Field:=14
    
'Adding the VLOOKUP Table
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("N5").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("N6").Select
    ActiveCell.FormulaR1C1 = "20"
    Range("O6").Select
    ActiveCell.FormulaR1C1 = "Amy"
    Range("O5").Select
    ActiveCell.FormulaR1C1 = "Patrick"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "Kirk"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = "Diane"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Amy"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Kevin"
    Range("AE7").Select
    ActiveCell.FormulaR1C1 = "Buyer"
    Range("AE8").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-26],R1C14:R6C15,2,FALSE)"
    
'Select range for pivot table
    Range("AE7").Select
    Selection.End(xlToLeft).Select
    Sheets.Add Before:=Sheets(1)
    Sheets(1).Name = "PO Pivot"

'Pivot table creation
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "OpenPO", Version:=xlPivotTableVersion10).CreatePivotTable TableDestination _
        :="Sheet1!R3C1", TableName:="PivotOpenPO", DefaultVersion:= _
        xlPivotTableVersion10
    Sheets("PO Pivot").Select
    Cells(3, 1).Select
    
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Buyer")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotOpenPO").AddDataField ActiveSheet.PivotTables( _
        "PivotOpenPO").PivotFields("CONFIRM CD"), "Count of CONFIRM CD", xlCount
    ActiveSheet.PivotTables("PivotOpenPO").AddDataField ActiveSheet.PivotTables( _
        "PivotOpenPO").PivotFields("CONF QTY"), "Count of CONF QTY", xlCount
    With ActiveSheet.PivotTables("PivotOpenPO").DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotOpenPO").AddDataField ActiveSheet.PivotTables( _
        "PivotOpenPO").PivotFields("CONF REC"), "Count of CONF REC", xlCount
    ActiveSheet.PivotTables("PivotOpenPO").AddDataField ActiveSheet.PivotTables( _
        "PivotOpenPO").PivotFields("BCK ORD"), "Count of BCK ORD", xlCount
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("VEND NAME")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("DATE DUE ")
        .Orientation = xlRowField
        .Position = 2
    End With
    
'No Subtotals
    ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Buyer").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotOpenPO").PivotFields("DATE DUE ").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotOpenPO").PivotFields("VEND NAME").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotOpenPO").TableStyle2 = "PivotStyleLight15"
    
'Change Counts to Sums
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Count of CONFIRM CD")
        .Caption = "CD"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Count of CONF QTY")
        .Caption = "QTY"
        .Function = xlSum
    End With
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Count of CONF REC")
        .Caption = "Sum of CONF REC"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Sum of CONF REC").Caption _
        = "Rec"
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("Count of BCK ORD")
        .Caption = "Bk Ord"
        .Function = xlSum
    End With
    
'Add PO # - should be up top
    With ActiveSheet.PivotTables("PivotOpenPO").PivotFields("PO #")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotOpenPO").PivotFields("PO #").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    
'Select All, autofit and align left
    Cells.Select
    Cells.EntireColumn.AutoFit
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
    Cells.EntireColumn.AutoFit
    Range("A2").Select
End Sub
