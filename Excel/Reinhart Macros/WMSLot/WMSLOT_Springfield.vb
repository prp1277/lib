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
        "Kevin Rosecrants"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "PENNY"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Todd Kamp"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet4").Select
    Sheets("Sheet4").Name = "TODD"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1, Criteria1:= _
        "Vincent Romi"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet5").Select
    Sheets("Sheet5").Name = "VINCENT"
    ActiveSheet.Paste
    Cells.Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Selection.ColumnWidth = 8.57
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveSheet.Range("$A$1:$AS$1053").AutoFilter Field:=1
    Sheets("ALL SFD").Select
    Range("A1").Select
    ActiveWorkbook.SaveCopyAs ("C:\Users\PRPowell\Desktop\NonShared\WMS Aging Files\Springfield\" & Format("120 Day Aging - Springfield mm.dd.yyyy") & ".xlsm")
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open ("C:\Users\PRPowell\Desktop\Nonshared\WMS Aging Files\Springfield\" & ("WMSLOT Springfield mm.dd.yyyy.xlsm"))
End Sub