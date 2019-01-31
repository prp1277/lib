Sub PivotData()
'
' PivotFormat Macro
' Formats the SQL Export into Pivot-Ready data
'

'
    Sheets("Invoices").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1)), TrailingMinusNumbers:=True
    Range("B5").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Invoices!R5C1:R1000C17", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="ALLERRORS", DefaultVersion _
        :=xlPivotTableVersion15
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "ALL ERRORS"
    Sheets("Macro").Select
    ActiveWorkbook.SaveCopyAs ("C:\Users\PRPowell\Desktop\Nonshared\Perfect Orders\Kansas City\" & Format("Perfect Order - KC mm.dd.yyy") & ".xlsm")
    ActiveWorkbook.Close SaveChanges:=False
    Workbooks.Open ("C:\Users\PRPowell\Desktop\Nonshared\Perfect Orders\Kansas City\" & ("Perfect Order - KC mm.dd.yyy"))
End Sub
----------------------------------------------------------------------------------------------------
Sub PivotData()		'AllDiv				***9/27/2017***
'
' PivotFormat Macro
' Formats the SQL Export into Pivot-Ready data
'

'
    Sheets("Invoices").Select
'This copies the =LEFT formula and pastes the values in (A:A)
    With Range("A5").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With

'This is the text to columns section using the =LEFT formula
    With Range("B6").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1)), TrailingMinusNumbers:=True
    End With
    Range("B5").Select

'Create the pivot table using Invoices(A5:Q1000) as source and Sheet1(A3) as the destination
'It also names the table, which will be referred to throughout the next few steps
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Invoices!R5C1:R1000C17", Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="ALLERRORS", DefaultVersion _
        :=xlPivotTableVersion15
    'xlPivotTableVersion15 = Excel 2013
    'xlPivotTableVersion14 = Excel 2010

    Sheets("Sheet1").Name = "ALL ERRORS"
    Sheets("Macro").Select
End Sub
