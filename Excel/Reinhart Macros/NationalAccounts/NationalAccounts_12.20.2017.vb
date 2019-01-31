Sub PCLessThan6()
'
' PCLessThan6 Macro
'

'
    ActiveSheet.Paste
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\FoxAndHound.xlsx"
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Fox&Hound"
    Application.WindowState = xlNormal
    Sheets("Fox&Hound").Select
    Application.CutCopyMode = False
    Sheets("Fox&Hound").Move Before:=Workbooks("Book1").Sheets(1)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Hallmark.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Hallmark"
    Sheets("Hallmark").Select
    Sheets("Hallmark").Move Before:=Workbooks("Book1").Sheets(2)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\MinitMart.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "MinitMart"
    Sheets("MinitMart").Select
    Sheets("MinitMart").Move Before:=Workbooks("Book1").Sheets(3)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Minskys.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Minskys"
    Sheets("Minskys").Select
    Sheets("Minskys").Move Before:=Workbooks("Book1").Sheets(4)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Noodles.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Noodles"
    Sheets("Noodles").Select
    Sheets("Noodles").Move Before:=Workbooks("Book1").Sheets(5)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Picklemanns.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Picklemanns"
    Sheets("Picklemanns").Select
    Sheets("Picklemanns").Move Before:=Workbooks("Book1").Sheets(6)
    Workbooks.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\RaisingCanes.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "RaisingCanes"
    Sheets("RaisingCanes").Select
    Sheets("RaisingCanes").Move Before:=Workbooks("Book1").Sheets(7)
    Workbooks.Open Filename:="I:\Purchasing\Reports\NationalAccounts\Sonic.xlsx"
    Application.WindowState = xlNormal
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Sonic"
    Sheets("Sonic").Select
    Sheets("Sonic").Move Before:=Workbooks("Book1").Sheets(8)
    Rows("1:9").Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("A2").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "KAN"
    Sheets(Array("Fox&Hound", "Hallmark", "MinitMart", "Minskys", "Noodles", _
        "Picklemanns", "RaisingCanes", "Sonic", "KAN")).Select
    Sheets("KAN").Activate
    Range("A8:H8").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("L8").Select
    Selection.Cut
    Range("L9").Select
    ActiveSheet.Paste
    Range("N8").Select
    Selection.Cut
    Range("N9").Select
    ActiveSheet.Paste
    Range("P8:Q8").Select
    Selection.Cut
    Range("P9").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("L38").Select
    Sheets("KAN").Select
    Range("A9").Select
    Sheets("Sonic").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$71"), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=1, Criteria1:= _
        "=KNX", Operator:=xlOr, Criteria2:="=SFD"
    Range("A32").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=1
    Range("A10").Select
    Sheets("RaisingCanes").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$48"), , xlYes).Name = _
        "Table2"
    Range("Table2[[#Headers],[DIV]]").Select
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1, Criteria1:= _
        Array("JNC", "MIL", "NBD", "NOR", "TID", "TWC"), Operator:=xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table2").Range.AutoFilter Field:=1
    Range("A10").Select
    Sheets("Picklemanns").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$11"), , xlYes).Name = _
        "Table3"
    Range("A10").Select
    Sheets("Noodles").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$49"), , xlYes).Name = _
        "Table4"
    Range("Table4[#All]").Select
    ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=1, Criteria1:= _
        Array("BOS", "DET", "JNC", "MIL", "PIT", "SHR", "SUN", "TWC"), Operator:= _
        xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=1
    Range("A10").Select
    Sheets("Minskys").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$15"), , xlYes).Name = _
        "Table5"
    Range("A10").Select
    Sheets("MinitMart").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$13"), , xlYes).Name = _
        "Table6"
    Range("Table6[#All]").Select
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1, Criteria1:= _
        "BGN"
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1
    Range("A10").Select
    Sheets("Hallmark").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$12"), , xlYes).Name = _
        "Table7"
    Range("A10").Select
    Sheets("Fox&Hound").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$50"), , xlYes).Name = _
        "Table8"
    Range("Table8[#All]").Select
    ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=1, Criteria1:= _
        Array("CIN", "EPA", "JNC", "LAX", "MIL", "SFD", "SHR", "TWC"), Operator:= _
        xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=1
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("KAN").Select
    Range("A10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$13"), , xlYes).Name = _
        "Table9"
    Range("A10").Select
    Selection.End(xlDown).Select
    Range("A14").Select
    Sheets("Hallmark").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    Range("A17").Select
    Sheets("MinitMart").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A20").Select
    Sheets("Minskys").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    Range("A26").Select
    Sheets("Noodles").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Range("A26").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    Range("A33").Select
    Sheets("Picklemanns").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A35").Select
    Sheets("RaisingCanes").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Sheets("KAN").Name = "KAN"
    Sheets("RaisingCanes").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    Range("A41").Select
    Sheets("Sonic").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select
    Range("A41").Activate
    Cells.EntireColumn.AutoFit
    Range("A3").Select
    Columns("A:A").ColumnWidth = 7.36
    Range("Table9[[#Headers],[DIV]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A8").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A8:Q9").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Columns("N:N").Select
    Range("N3").Activate
    Selection.NumberFormat = "m/d/yyyy"
    Columns("B:B").Select
    Range("B3").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Range("Q8").Select
    ActiveCell.FormulaR1C1 = "OH +"
    Range("Q9").Select
    ActiveCell.FormulaR1C1 = "OO"
    Range("Q10").Select
    ActiveCell.FormulaR1C1 = "=[@HAND]+[@ORDER]"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Sum /"
    Range("R9").Select
    ActiveCell.FormulaR1C1 = "Avg Move"
    Range("R10").Select
    ActiveCell.FormulaR1C1 = "=[@OO]/[@MOVEMENT]"
    Range("Table9[Avg Move]").Select
    Range("R11").Activate
    Selection.NumberFormat = "#,##0.00"
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Weeks"
    Range("Table9[[#Headers],[Avg Move]]").Select
    ActiveCell.FormulaR1C1 = "Ordered"
    Range("S9").Select
    ActiveCell.FormulaR1C1 = "Comment"
    Range("S10").Select
    ActiveCell.FormulaR1C1 = "=IF([@OO]<1,""Potential Out of Stock"","""")"
    Range("Q8:S9").Select
    Range("Table9[[#Headers],[OO]]").Activate
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.Select
    Range("E3").Activate
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Columns("A:A").ColumnWidth = 6.55
    Cells.Select
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
    Range("C10").Select
    ActiveWindow.FreezePanes = True
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("E1").Select
    Selection.Copy
    Application.CutCopyMode = False
End Sub
