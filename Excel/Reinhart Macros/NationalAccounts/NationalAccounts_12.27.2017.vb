Sub NationalAccountsImport()	'12/27/2017
'
' Macro1 Macro
'
    Application.ScreenUpdating = False
    
'Fox and Hound Hyperlink in A2
    Range("A2").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Fox&Hound"
    Sheets("Fox&Hound").Select
    Sheets("Fox&Hound").Move After:=Workbooks("NationalAccounts1").Sheets(1)
    Sheets("Sheet1").Select
    
'Hallmark Hyperlink in A3
    Range("A3").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Hallmark"
    Sheets("Hallmark").Select
    Sheets("Hallmark").Move After:=Workbooks("NationalAccounts1").Sheets(2)
    Sheets("Sheet1").Select

'Minit Mart Hyperlink in A4
    Range("A4").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "MinitMart"
    Windows("NationalAccounts1").Activate
    Windows("MinitMart.xlsx").Activate
    Sheets("MinitMart").Select
    Sheets("MinitMart").Move After:=Workbooks("NationalAccounts1").Sheets(3)
    
'Minskys Hyperlink in A5
    Sheets("Sheet1").Select
    Range("A5").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Minskys"
    Sheets("Minskys").Select
    Windows("NationalAccounts1").Activate
    Windows("Minskys.xlsx").Activate
    Sheets("Minskys").Select
    Sheets("Minskys").Move After:=Workbooks("NationalAccounts1").Sheets(4)
    Sheets("Sheet1").Select
    
'Noodles Hyperlink in A6
    Range("A6").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Noodles"
    Sheets("Noodles").Select
    Sheets("Noodles").Move After:=Workbooks("NationalAccounts1").Sheets(5)
    
'Picklemanns Hyperlink in A7
    Sheets("Sheet1").Select
    Range("A7").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Picklemanns"
    Sheets("Picklemanns").Select
    Sheets("Picklemanns").Move After:=Workbooks("NationalAccounts1").Sheets(6)
    
'Raising Canes Hyperlink in A8
    Sheets("Sheet1").Select
    Range("A8").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "RaisingCanes"
    Sheets("RaisingCanes").Select
    Sheets("RaisingCanes").Move After:=Workbooks("NationalAccounts1").Sheets(7)
    Sheets("Sheet1").Select
    
'Sonic Hyperlink in A9
    Range("A9").Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Sonic"
    Windows("NationalAccounts1").Activate
    Windows("Sonic.xlsx").Activate
    Sheets("Sonic").Select
    Sheets("Sonic").Move After:=Workbooks("NationalAccounts1").Sheets(8)
    Sheets("Sheet1").Select

End Sub
-------------------------------------------------------------------------------
Sub AbsoluteRef()
'
' AbsoluteRef Macro
' Same has "Headers" macro, but absolute references
'

'
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
    Range("A9").Select
End Sub
-----------------------------------------------------------------------------------------------
Sub Macro3()
'
' Macro3 Macro
' Format as table, filter out KAN and delete rows. Then, clear filter so KAN is the only division left
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$9:$Q$52"), , xlYes).Name = _
        "Table6"
    Range("Table6[#All]").Select
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1, Criteria1:= _
        Array("CIN", "EPA", "JNC", "LAX", "MIL", "SFD", "SHR", "TWC"), Operator:= _
        xlFilterValues
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Table6").Range.AutoFilter Field:=1
    Range("A11").Select
End Sub