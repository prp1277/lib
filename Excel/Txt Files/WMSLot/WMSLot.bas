Attribute VB_Name = "WMSLot"
                                                                                                                                                       
                                                                                                                                                       Sub HundredTwentyWMSLotQueryPrep()
'
' HundredTwentyQueryPrep Macro
' Prepare the HundredTwenty WMSLot file to be imported to a query
'

    Application.ScreenUpdating = False
    Range("HundredTwenty[[#All],[DIV]:[Brand]]").Select
    Range("HundredTwenty[[#Headers],[Brand]]").Activate
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Range("HundredTwenty[[#All],[Description]:[CW QTY]]").Select
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Range("HundredTwenty[[#All],[WeeklyMove]:[Wk Onh]]").Select
    Selection.ListObject.ListColumns(5).Delete
    Selection.ListObject.ListColumns(5).Delete
    Range("HundredTwenty[[#All],[Shelf Life]:[Rec Date]]").Select
    Selection.ListObject.ListColumns(6).Delete
    Selection.ListObject.ListColumns(6).Delete
    Range("HundredTwenty[[#All],[Check Life]:[Shlf Verif]]").Select
    Selection.ListObject.ListColumns(8).Delete
    Selection.ListObject.ListColumns(8).Delete
    Range("HundredTwenty[[#All],[License]:[PO '#]]").Select
    Selection.ListObject.ListColumns(12).Delete
    Selection.ListObject.ListColumns(12).Delete
    Range("HundredTwenty[[#Headers],[User Comments]]").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("HundredTwenty[#All]").Select
    Range("HundredTwenty[[#Headers],[User Comments]]").Activate
    Selection.Columns.AutoFit
    Range("HundredTwenty[[#All],[Pick Slot]]").Select
    Selection.ListObject.ListColumns(13).Delete
    Range("HundredTwenty[[#Headers],[Tot Reserve]]").Select
    Selection.End(xlToLeft).Select
    Range("HundredTwenty[[#Headers],[Prod'#]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.End(xlToLeft).Select
    Range("HundredTwenty[[#All],[LotCst]]").Select
    Selection.ListObject.ListColumns(3).Delete
    Range("HundredTwenty[[#All],[Lot Trkd]:[LT Exp Date]]").Select
    Selection.ListObject.ListColumns(5).Delete
    Selection.ListObject.ListColumns(5).Delete
    Range("HundredTwenty[[#Headers],[Prod'#]]").Select
End Sub

Sub WMSLotDailyQueryPrep()
'
' DailyQueryPrep Macro
' Prepare the daily WMSLot file to be imported to a query
'

'
    Range("Daily[[#All],[DIV]:[Brand]]").Select
    Range("Daily[[#Headers],[Brand]]").Activate


'Use columns formula to see what 12 columns are deleted here
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Range("Daily[[#All],[Description]:[CW QTY]]").Select


'Use columns formula to see what 12 columns are deleted here
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Selection.ListObject.ListColumns(3).Delete
    Range("Daily[[#All],[WeeklyMove]:[Wk Onh]]").Select
    Selection.ListObject.ListColumns(5).Delete
    Selection.ListObject.ListColumns(5).Delete
    Range("Daily[[#All],[Shelf Life]:[Rec Date]]").Select
    Selection.ListObject.ListColumns(6).Delete
    Selection.ListObject.ListColumns(6).Delete
    Range("Daily[[#All],[Check Life]:[Shlf Verif]]").Select
    Selection.ListObject.ListColumns(8).Delete
    Selection.ListObject.ListColumns(8).Delete
    Range("Daily[[#All],[License]:[PO '#]]").Select
    Selection.ListObject.ListColumns(12).Delete
    Selection.ListObject.ListColumns(12).Delete
    Range("Daily[[#Headers],[User Comments]]").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Daily[#All]").Select
    Range("Daily[[#Headers],[User Comments]]").Activate
    Selection.Columns.AutoFit
    Range("Daily[[#All],[Pick Slot]]").Select
    Selection.ListObject.ListColumns(13).Delete
    Range("Daily[[#Headers],[Tot Reserve]]").Select
    Selection.End(xlToLeft).Select
    Range("Daily[[#Headers],[Prod'#]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.End(xlToLeft).Select
    Range("Daily[[#All],[LotCst]]").Select
    Selection.ListObject.ListColumns(3).Delete
    Range("Daily[[#All],[Lot Trkd]:[LT Exp Date]]").Select
    Selection.ListObject.ListColumns(5).Delete
    Selection.ListObject.ListColumns(5).Delete
    Range("Daily[[#Headers],[Prod'#]]").Select
End Sub

