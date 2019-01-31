Attribute VB_Name = "PerfectOrder"
Sub CreatePivotCache()
'
' PivotFormat Macro
' Formats the SQL Export into Pivot-Ready data
'
    Dim Sh As Sheets
    Set Sh = Sheets
    
'
    Sh("Invoices").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("B6"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(5, 1)), TrailingMinusNumbers:=True
    Range("B5").Select
    Sh.Add Before:=Sh(1)
    Sh(1).Name = "Sheet1"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Invoices!R5C1:R1000C17") _
        .CreatePivotTable _
        TableDestination:="Sheet1!R3C1", TableName:="ALLERRORS"
    Sh("Sheet1").Select
    Sh("Sheet1").Name = "ALL ERRORS"
    Sheets("Macro").Select
End Sub

Sub AllErrors()
'
' ALLERRORS Macro
' Creates the ALLERRORS Table
'
    Dim Sh As Sheets
    Set Sh = Sheets
    
    Sh("ALL ERRORS").Select

'You have to go in order of the columns, otherwise the pivot table gets mad
'Format A#
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 1
        .PivotItems("").Visible = False
    End With

'Format Customer
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'Format Invoice Number
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice  #")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'Format Invoice Date
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'Format Responsible
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 5
        .Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
 
'Format L1 Error
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 6
    End With
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With


'Add the count of L1 Errors Column
    ActiveSheet.PivotTables("ALLERRORS").AddDataField ActiveSheet.PivotTables( _
        "ALLERRORS").PivotFields("L1 Error"), "Count of L1 Error", xlCount
    With ActiveSheet.PivotTables("ALLERRORS").PivotFields("L1 Error")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
'Format as a tabular pivot table
    With ActiveSheet.PivotTables("ALLERRORS")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    Range("A2").Select
    
'End and autofit columns
    Cells.EntireColumn.AutoFit
    Sh("Macro").Select
End Sub

Sub ExecutionErrors()
'
' ExecutionErrorPivot Macro
' Creates the Execution Error Pivot Table
'
'Declarations
    Dim Sh As Sheets
    Set Sh = Sheets

'Add A new sheet and name it Sheet 2
    Sh.Add Before:=Sh(1)
    Sh(1).Name = "Sheet2"
    
'Create the new pivot to Sheet2 then rename it to Execution Errors
    Sh("ALL ERRORS").PivotTables("ALLERRORS"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet2!R3C1", TableName:= _
        "EXECUTIONERRORS"
    Sh(1).Select
    Sh(1).Name = "Execution Errors"
    Cells(3, 1).Select
    Sh("Execution Errors").Select
    
'Format Customer
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'Format A#
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("A#")
        .Orientation = xlRowField
        .Position = 2
    End With

'Format Responsible
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
'Format Invoice Date
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
'Set L1 as a filter and select only Execution Errors
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .PivotItems("Availability Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    
'Format L1 Error
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'Format Product #
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 6
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
'Set L1 as a filter and select only Execution Errors
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .PivotItems("Availability Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    
'Format L3 Errors
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 6
        .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
'Format L1 Error
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    
'Format the count of L3 Errors
    ActiveSheet.PivotTables("EXECUTIONERRORS").AddDataField ActiveSheet.PivotTables( _
        "EXECUTIONERRORS").PivotFields("L3 Error"), "Count of L3 Error", xlCount
    With ActiveSheet.PivotTables("EXECUTIONERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 7
    End With
    
'Format Tabular Pivot Table
    With ActiveSheet.PivotTables("ExecutionErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With
    
'End by autofitting and going back to the macros tab
    Cells.EntireColumn.AutoFit
    Sheets("Macro").Select
End Sub

Sub AvailabilityErrors()
'
' AvailabilityErrors Macro
' Creates the Availability Errors Tab and Pivot Table
'
    Dim Sh As Sheets
    Set Sh = Sheets
    Application.ScreenUpdating = False
    
'Add a new sheet and name it Sheet3
    Sh.Add Before:=Sh(1)
    Sh(1).Name = "Sheet3"
    
'Create new PivotTable and name it Availability Errors
    Sh("ALL ERRORS").PivotTables("ALLERRORS"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet3!R3C1", TableName:= _
        "AVAILABILITYERRORS"
    Sh(1).Select
    Sh(1).Name = "Availability Errors"
    
 'Format Responsible
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Responsible")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With
    
'Format Customer
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With
    
'Format Product #
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format Invoice Date
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Add the Count of L3 Errors datafield
    ActiveSheet.PivotTables("AVAILABILITYERRORS").AddDataField ActiveSheet.PivotTables( _
        "AVAILABILITYERRORS").PivotFields("L3 Error"), "Count of L3 Error", xlCount

'Format L3 Error
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With
    
'Set L1 Error as the filter
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("AVAILABILITYERRORS").PivotFields("L1 Error")
        .PivotItems("Execution Error").Visible = False
        .PivotItems("Order Entry Error").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    
'Set the formatting to old school pivot table
    With ActiveSheet.PivotTables("AvailabilityErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With

'Autowidth and back to Macro tab
    Cells.EntireColumn.AutoFit
    Sheets("Macro").Select
End Sub

Sub ProductDetails()
'Creates the Product Details Pivot Table

'Declarations
    Dim Sh As Sheets
    Set Sh = Sheets
    Application.ScreenUpdating = False
    
'Add a new sheet and name it Sheet4
    Sh.Add Before:=Sh(1)
    Sh(1).Name = "Sheet4"
    
'Create the Pivot table and name it ProductDetails
    Sh("ALL ERRORS").PivotTables("ALLERRORS"). _
        PivotCache.CreatePivotTable TableDestination:="Sheet4!R3C1", TableName:= _
        "ProductDetails"
    Sh(1).Select
    Sh(1).Name = "Product Details"
    
'Format Description
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Description")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format Product #
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 2
    End With

'Format Invoice Date
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Invoice Date")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format L2 Error
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("L2 Error")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format L3 Errors
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("L3 Error")
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format Vendor Name
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Vendor Name")
        .Orientation = xlRowField
        .Position = 6
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Format Customer
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Customer")
        .Orientation = xlRowField
        .Position = 7
        .Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, _
            False, False)
    End With

'Add the Count of Product # datafield
    ActiveSheet.PivotTables("ProductDetails").AddDataField ActiveSheet.PivotTables( _
        "ProductDetails").PivotFields("Product #"), "Count of Product #", xlCount
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("Product #")
        .Orientation = xlRowField
        .Position = 8
    End With
    
'Set L1 Error as the filter
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("L1 Error")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("ProductDetails").PivotFields("L1 Error")
        .PivotItems("(blank)").Visible = False
    End With
    
'Set to tabular formatting
    With ActiveSheet.PivotTables("AvailabilityErrors")
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
    End With

'Auto width and back to Macro tab
    Cells.EntireColumn.AutoFit
    Sheets("Macro").Select
    
End Sub


