Attribute VB_Name = "NationalAccounts"
Sub NTLAccImport()
'Import sheets from I:\ location, name it and place it
'At the beginning of the workbook
    
    Dim Wb As Workbook
    Set Wb = Workbooks
    
    Dim Sh As Sheets
    Set Sh = Sheets
    
'Fox & Hound Import
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\FoxAndHound.xlsx"
    Sh(1).Select
    With Sh("Sheet1").Name = "Fox&Hound"
        .Move Before:=Workbooks("Book1").Sheets(1)
    End With

'Hallmark Import
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Hallmark.xlsx"
    With Sh(1).Select
        .Name = "Hallmark"
        .Move Before:=Workbooks("Book1").Sheets(2)
    End With
    
'Import Minit Mart
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\MinitMart.xlsx"
    With Sh(1).Select
        .Name = "MinitMart"
        .Move Before:=Workbooks("Book1").Sheets(3)
    End With
    
'Import Minskys
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Minskys.xlsx"
    With Sh(1).Select
        .Name = "Minskys"
        .Move Before:=Workbooks("Book1").Sheets(4)
    End With
    
'Import Noodles
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Noodles.xlsx"
    With Sh(1).Select
        .Name = "Noodles"
        .Move Before:=Workbooks("Book1").Sheets(5)
    End With
    
'Import Picklemanns
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\Picklemanns.xlsx"
    With Sh(1).Select
        .Name = "Picklemanns"
        .Move Before:=Workbooks("Book1").Sheets(6)
    End With
    
'Import Raising Canes
    Wb.Open Filename:= _
        "I:\Purchasing\Reports\NationalAccounts\RaisingCanes.xlsx"
    With Sh(1).Select
        .Name = "RaisingCanes"
        .Move Before:=Workbooks("Book1").Sheets(7)
    End With
    
'Import Sonic
    Wb.Open Filename:="I:\Purchasing\Reports\NationalAccounts\Sonic.xlsx"
    With Sh(1).Select
        .Name = "Sonic"
        .Move Before:=Workbooks("Book1").Sheets(8)
    End With
End Sub
Sub PCTables()
'Copy Report headers, name KAN tab and format Tables


'Copy First 9 Rows from Sonic and Paste to KAN
    Sheets(8).Activate
    Rows("1:9").Select
    Selection.Copy
    Sheets(9).Select
    Range("A1").Select
    ActiveSheet.Paste
    
'Name KAN Tab
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "KAN"
    
'Table Headers Cut/Paste
    Range("A8:H8").Select
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
    
    'Table1 - Sonic
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
    
    'Table2 - Canes
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
    
    'Table3 - Picklemanns - very clean example
    Sheets("Picklemanns").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$11"), , xlYes).Name = _
        "Table3"
    Range("A10").Select
    
    'Table4 - Noodles
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
    
    'Table5 - Minskys
    Sheets("Minskys").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$15"), , xlYes).Name = _
        "Table5"
    Range("A10").Select
    
    'Table6 - MinitMart
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
    
    'Table7 - Hallmark
    Sheets("Hallmark").Select
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$12"), , xlYes).Name = _
        "Table7"
    Range("A10").Select
    
    'Table8 - Fox & Hound
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
    
    'Table9 - KAN
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
    
End Sub

Sub CopyToKAN()
'Copy pages to KAN
    'Hallmark - A17 End Range
    Sheets("Hallmark").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlDown).Select
    Range("A17").Select
    
    'Minit Mart - A20 End Range
    Sheets("MinitMart").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A20").Select
    
    'Minskys - A26 End Range
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
    
    'Noodles - A33 End Range
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
    
    'Picklemanns - A35 End Location
    Sheets("Picklemanns").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("KAN").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A35").Select
    
    'Raising Canes - A41 End Location
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
    
    'Sonic - A41 End Location
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
    
End Sub

Sub PCFormatandFormulas()
    
    'Paste Format
    Cells.EntireColumn.AutoFit
    Range("A3").Select
    Columns("A:A").ColumnWidth = 7.36
    Range("Table9[[#Headers],[DIV]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("A8").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    'White Text Headers
    Range("A8:Q9").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    'Format date
    Columns("N:N").Select
    Selection.NumberFormat = "m/d/yyyy"
    
    'Delete Div # Column
    Columns("B:B").Select
    Range("B3").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    
    'OH + OO Formula
    Range("Q8").Select
    ActiveCell.FormulaR1C1 = "OH +"
    Range("Q9").Select
    ActiveCell.FormulaR1C1 = "OO"
    Range("Q10").Select
    ActiveCell.FormulaR1C1 = "=[@HAND]+[@ORDER]"
    
    'Sum / Avg Move Formula
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Sum /"
    Range("R9").Select
    ActiveCell.FormulaR1C1 = "Avg Move"
    Range("R10").Select
    ActiveCell.FormulaR1C1 = "=[@OO]/[@MOVEMENT]"
    
    'Avg Move # Format
    Range("Table9[Avg Move]").Select
    Range("R11").Activate
    Selection.NumberFormat = "#,##0.00"
    
    'Weeks Ordered
    Range("R8").Select
    ActiveCell.FormulaR1C1 = "Weeks"
    Range("Table9[[#Headers],[Avg Move]]").Select
    ActiveCell.FormulaR1C1 = "Ordered"
    
    'Comment Formula
    Range("S9").Select
    ActiveCell.FormulaR1C1 = "Comment"
    Range("S10").Select
    ActiveCell.FormulaR1C1 = "=IF([@OO]<1,""Potential Out of Stock"","""")"
    
    'Table Formatting
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
    
    'Cell Formatting
    Cells.Select
    Range("E3").Activate
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Columns("A:A").ColumnWidth = 6.55
    Cells.Select
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
    
    'Freeze Panes
    Range("C10").Select
    ActiveWindow.FreezePanes = True
    
    'TODAY() formula - Date copied to clipboard
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("E1").Select
    Selection.Copy
    Application.CutCopyMode = False
    
End Sub

