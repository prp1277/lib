Sub ErrorsPerCustomer()
'
' ErrorsPerCustomer Macro
' Create the Errors Per Customer Tab
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Errors Per Customer"
    Sheets("Invoices").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Errors Per Customer").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "A#"
    Range("A2").Select
    Sheets("Invoices").Select
    Selection.Copy
    Sheets("Errors Per Customer").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Delete Everything up until this point"
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$A$997").RemoveDuplicates Columns:=1, Header:=xlYes
    Range("A1").Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Customer"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1]&"" ""&R1C3,'ALL ERRORS'!R3C2:R894C9,7,FALSE)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B351")
    Range("B2:B351").Select
    Sheets("Errors Per Customer").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1]&"" ""&R1C3,'ALL ERRORS'!R3C2:R894C9,8,FALSE)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B351")
    Range("B2:B351").Select
    Selection.Cut Destination:=Range("C2:C351")
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Invoices!R5C1:R670C2,2,FALSE)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B351")
    Range("B2:B351").Select
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("C:C").Select
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
    Range("C1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("C2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A2:C350").Select
    Range("C2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlLTR
        .MergeCells = False
    End With
    Range("A1:C1").Select
    Range("C1").Activate
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Errors Per Customer").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Errors Per Customer").AutoFilter.Sort.SortFields. _
        Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Errors Per Customer").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:R[348]C[-2])"
    Range("E3").Select
    ActiveWindow.SmallScroll Down:=-24
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C[-2]:R[347]C[-2])"
    Range("E4").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "More Efficient"
    Range("F4").Select
End Sub