Sub BuyerTabsSFD()
'
' BuyerTabsSFD Macro
' Copy and Paste the WMSLOT Data to be separated by buyer
'

'
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "ALL KAN"
    Columns("A:A").Select
    Range("A2").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Amy Davis"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "AMY"
    Sheets("ALL KAN").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("AMY").Select
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Diane Cole"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "DIANE"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Kevin Rosencrants"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "Kevin"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Kirk Thompson"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "KIRK"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Patrick Powell"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "PATRICK"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Selection.ColumnWidth = 8.57
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1
    Sheets("ALL KAN").Select
    Range("A1").Select
    ActiveWorkbook.SaveCopyAs ("C:\Users\FWWhite\Desktop\" & Format("WMSLOT - KAN mm.dd.yyyy") & ".xlsm")
    Sheets(
End Sub

Sub WMSLot120PropCodes

    Sheets("KAN").Select
    Range("A1").Select
    Sheets("KAN").Range("A1:AS").

-----------------------------------------------------------------------------------------------------------------
Sub PropCodes()
'
' PropCodes Macro
' Separate tabs by prop codes
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "J1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$AU$807").AutoFilter Field:=10, Criteria1:="WP"
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "WP - PepperJax"
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("A1").Select
    ActiveWorkbook.Worksheets("WP - PepperJax").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("WP - PepperJax").AutoFilter.Sort.SortFields.Add Key _
        :=Range("C1:C64739"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("WP - PepperJax").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub
-----------------------------------------------------------------------------------------------------------------

Sub PropCodes2()
'
' PropCodes2 Macro
' Shorter Version
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AU$807").AutoFilter Field:=10, Criteria1:="VD"
    ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "C1:C807"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("KAN").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1:AU807").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "VD - Sonic"
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("A1").Select
End Sub