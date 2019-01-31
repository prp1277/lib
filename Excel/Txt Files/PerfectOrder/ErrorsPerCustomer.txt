Sub ErrorPerCustomer()
'
' ErrorPerCust Macro
' Creates the Error/Customer Tab
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet3").Select
    Sheets("Sheet3").Name = "# Issues By Customer"
    ActiveCell.FormulaR1C1 = "A#"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "LOOKUP VALUE"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Customer"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B2").Select
    Sheets("Invoices").Select
    Range("A704").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A6:A704").Select
    Range("A704").Activate
    Selection.Copy
    Sheets("# Issues By Customer").Select
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$A$700").RemoveDuplicates Columns:=1, Header:=xlYes
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-1],"" "",R1C4)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B329")
    Range("B2:B329").Select
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'ALLERRORS'!R[1]C[-2]:R[538]C[6],2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R[1]C[-2]:R[538]C[6],2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R3C1:R540C9,2,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C329")
    Range("C2:C329").Select
    Range("D2").Select
    Sheets("ALLERRORS").Select
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
    End With
    Sheets("# Issues By Customer").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'ALLERRORS'!R3C1:R540C9,9,FALSE)"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D329")
    Range("D2:D329").Select

End Sub