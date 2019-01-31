Sub HeadersandConsolidate()
'
' HeadersandConsolidate Macro
' Shift the headers down one row, delete everyone but KAN and add all the data to the new sheet
'

'
    Sheets(Array("Fox&Hound", "Hallmark", "MinitMart", "Minsky's", "Noodles", _
        "Picklemanns", "Raising Canes", "Sonic", "KAN")).Select
    Sheets("Fox&Hound").Activate
    Range("A8:H8").Select
    Selection.Cut Destination:=Range("A9:H9")
    Range("L8").Select
    Selection.Cut Destination:=Range("L9")
    Range("N8").Select
    Range("N8").Cut Destination:=Range("N9")
    Range("P8:Q8").Select
    Selection.Cut Destination:=Range("P9:Q9")
    Range("P9:Q9").Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Sheets("MinitMart").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Fox&Hound").Select
    Range("C24").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$9:$Q$55"), , xlYes).Name = _
        "Table1"
    Range("A10").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$55"), , xlYes).Name = _
        "Table2"
    Range("A9:Q55").Select
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1, Criteria1:= _
        Array("CIN", "EPA", "JNC", "LAX", "MIL", "SFD", "SHR", "TWC"), Operator:= _
        xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Selection.End(xlToLeft).Select
    Range("A9").Select
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1
    ActiveWindow.SmallScroll Down:=-18
    Sheets("Hallmark").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$10"), , xlYes).Name = _
        "Table3"
    Range("A9").Select
    Sheets("MinitMart").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$11"), , xlYes).Name = _
        "Table4"
    Range("A9").Select
    ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=1, Criteria1:= _
        "BGN"
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=1
    Sheets("Minsky's").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$13"), , xlYes).Name = _
        "Table5"
    Range("A10").Select
    Sheets("Noodles").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$52"), , xlYes).Name = _
        "Table6"
    Range("A9:Q52").Select
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1, Criteria1:= _
        Array("BOS", "DET", "JNC", "MIL", "SUN", "TWC"), Operator:=xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Selection.End(xlToLeft).Select
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1
    ActiveWindow.SmallScroll Down:=-30
    Sheets("Picklemanns").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$13"), , xlYes).Name = _
        "Table7"
    Range("A9:Q13").Select
    ActiveSheet.ListObjects("Table7").Range.AutoFilter Field:=1, Criteria1:= _
        "BOS"
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table7").Range.AutoFilter Field:=1
    ActiveWindow.SmallScroll Down:=-3
    Sheets("Raising Canes").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$53"), , xlYes).Name = _
        "Table8"
    Range("A9:Q53").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("A10").Select
    ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=1, Criteria1:= _
        Array("JNC", "MIL", "NBD", "NOR", "TID", "TWC"), Operator:=xlFilterValues
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveWindow.ScrollColumn = 1
    ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=1
    ActiveWindow.SmallScroll Down:=-30
    Sheets("Sonic").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$81"), , xlYes).Name = _
        "Table9"
    Range("A9").Select
    ActiveSheet.ListObjects("Table9").Range.AutoFilter Field:=1, Criteria1:= _
        "=KNX", Operator:=xlOr, Criteria2:="=SFD"
    Range("A34").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table9").Range.AutoFilter Field:=1
    Range("A9").Select
    ActiveWindow.TabRatio = 0.761
    ActiveWindow.SmallScroll Down:=-9
    Sheets("KAN").Select
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "PROPRIETARY LISTING - ALL"
    Range("A3").Select
    Sheets("Sonic").Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveSheet.Next.Select
    Range("A10").Select
    ActiveSheet.Paste
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$33"), , xlYes).Name = _
        "Table10"
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    Sheets("Sonic").Select
    Range("C9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C9").Select
    ActiveSheet.Previous.Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Selection.End(xlToLeft).Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveSheet.Previous.Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    Sheets("KAN").Select
    ActiveWindow.SmallScroll Down:=0
    Sheets("Raising Canes").Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("KAN").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A34").Select
    ActiveSheet.Paste
    ActiveWindow.NewWindow
    Application.Left = 1328.5
    Application.Top = 74.5
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets("KAN").Select
    ActiveWindow.SmallScroll Down:=24
    ActiveWindow.Zoom = 85
    ActiveWindow.Zoom = 70
    ActiveWindow.SmallScroll Down:=15
    Windows("prop_VD7_COR121317.xls:1").Activate
    Sheets("Picklemanns").Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A38").Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Range("A41").Select
    Windows("prop_VD7_COR121317.xls:1").Activate
    ActiveSheet.Previous.Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Range("A49").Select
    Windows("prop_VD7_COR121317.xls:1").Activate
    ActiveSheet.Previous.Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Range("A53").Select
    Windows("prop_VD7_COR121317.xls:1").Activate
    ActiveSheet.Previous.Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A54").Select
    Windows("prop_VD7_COR121317.xls:1").Activate
    ActiveSheet.Previous.Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    ActiveSheet.Paste
    Range("A55").Select
    Windows("prop_VD7_COR121317.xls:1").Activate
    ActiveSheet.Previous.Select
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("prop_VD7_COR121317.xls:2").Activate
    ActiveSheet.Paste
    Range("A10:Q59").Select
    Range("A55").Activate
    Selection.Columns.AutoFit
    ActiveWindow.SmallScroll Down:=-81
    Cells.Select
    Application.CutCopyMode = False
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
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "OH + "
    Range("R9").Select
    ActiveCell.FormulaR1C1 = "OO"
    Range("S8").Select
    ActiveCell.FormulaR1C1 = "SUM / "
    Range("S9").Select
    ActiveCell.FormulaR1C1 = "Avg Wk"
    Range("T8").Select
    ActiveCell.FormulaR1C1 = ""
    Range("T9").Select
    ActiveCell.FormulaR1C1 = "Comment"
    Range("U9").Select
    ActiveCell.FormulaR1C1 = "Days Out"
    Range("R10").Select
    Columns("J:J").EntireColumn.AutoFit
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A8").Select
    Columns("A:A").ColumnWidth = 8.86
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    Range("R10").Select
    ActiveCell.FormulaR1C1 = "=RC[-8]+RC[-5]"
    Range("S10").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-4]"
    Range("T10").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<1,""Potential Out"","""")"
    Range("U10").Select
    ActiveCell.FormulaR1C1 = "=RC[-7]-TODAY()"
    Range("U10").Select
    Selection.Copy
    Application.CutCopyMode = False
    Range("U10:U59").Select
    Selection.NumberFormat = "d-mmm-yy"
    Selection.NumberFormat = "#,##0.00"
    ActiveSheet.ListObjects("Table10").Range.AutoFilter Field:=21, Criteria1:= _
        "#VALUE!"
    Selection.ClearContents
    ActiveSheet.ListObjects("Table10").Range.AutoFilter Field:=21
    ActiveWindow.SmallScroll Down:=-12
    Range("R9:U9").Select
    Selection.Copy
    Range("R8").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A8").Select
End Sub
