 Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Windows("WMSLOT_Macro.xlsm").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("O:T").Select
    Selection.Delete Shift:=xlToLeft
    Range("O1").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.AutoFilter
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub