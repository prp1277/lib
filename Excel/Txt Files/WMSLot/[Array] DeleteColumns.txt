Sub KevinDailyTest()

'This is the Daily WMSLot Macro for Kev
'BuyersTabKAN Macro
'This macro formats the WMSLOT Report into each buyer's tab
Application.ScreenUpdating = False
    

    Sheets(2).Select
    Sheets(2).Name = "Kevin"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Cells.EntireColumn.AutoFit
    
'As this deletes columns it's going to shift the others
'and mess up their positioning
'Step through and write one line at a time
    Application.DisplayAlerts = False
    Range("A:B").Delete
    Range("B:C").Delete
    Range("C:C").Delete
    Range("D:H").Delete
    Range("E:G").Delete
    Range("F:F").Delete
    Range("I:M").Delete
    Range("K:L").Delete
    Range("M:M").Delete
    Range("N:S").Delete
        
End Sub