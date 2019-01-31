Sub SalesCustomerTabs()
'
' SalesCustomerTabs Macro
' Takes the sales data and seperates into customer and sales tables
'
    Application.ScreenUpdating = False
    
    ActiveCell.Select
    Application.Run "Personal.xlsb!ActiveCellToTable"
    ActiveSheet.ListObjects("Table1").Name = "Sales"
    
'Copy the customer information to sheet 1 and name it customers
    Columns("C:I").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Range("A1").Select
    ActiveCell.Select
    Application.Run "Personal.xlsb!ActiveCellToTable"
    ActiveSheet.ListObjects("Table1").Name = "Customer"
    'Range(Selection, Selection.End(xlToRight)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    'ActiveSheet.ListObjects.Add(xlSrcRange, SLT, , xlYes).Name _
    '    = "Table2"
    'Range("Table2[#All]").Select
    'ActiveSheet.ListObjects("Table2").Name = "Customer"
    
'Delete the information that was just pasted into the customer tab
    Sheets("Results").Select
    Range("Sales[[#Headers],[Customer Name]:[Zip Code]]").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ListObject.ListColumns(4).Delete
    Selection.ListObject.ListColumns(4).Delete
    Selection.ListObject.ListColumns(4).Delete
    Selection.ListObject.ListColumns(4).Delete
    Selection.ListObject.ListColumns(4).Delete
    
'Autofit and format the Sales data
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
    Cells.EntireColumn.AutoFit
    Range("Sales[[#Headers],[Pack Size]]").Select
    Columns("G:G").ColumnWidth = 25.86
    Columns("H:H").ColumnWidth = 11.57
    Columns("R:R").ColumnWidth = 9.71
    Columns("T:T").ColumnWidth = 12.14
    Columns("E:E").ColumnWidth = 17.29
    
'Back to the customer tab to format and add the full address
    Sheets("Sheet1").Select
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Full Address"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=TRIM([@[Address One]]&"" ""&[@[Address Two]]&"" ""&[@City]&"" ""&[@[Zip Code]])"
    Range("H3").Select
    Columns("H:H").EntireColumn.AutoFit
    
'This is where I had to delete the product information after removing duplicates the first time through
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    
'This is removing the duplicate customers so they only show up once
    'Range(Selection, Selection.End(xlToRight)).Select
    'Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Range("Customer[#All]").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5 _
        , 6, 7), Header:=xlYes
    Cells.Select
    Cells.EntireColumn.AutoFit
    
'Back to the Results page
    Sheets("Results").Select
    Range("A2").Select

'Used the finder window to make sure I was going to the right place, then save with macros enabled
'ChDir "I:\Purchasing\Reports\PerfectOrders\KansasCity\QueryReference"
    ActiveWorkbook.SaveAs Filename:= _
        "I:\Purchasing\Reports\PerfectOrders\KansasCity\QueryReference\POSales.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub

