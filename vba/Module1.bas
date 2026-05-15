Attribute VB_Name = "Module1"
Dim POSheet As Worksheet
Option Explicit

' --- RefreshAll ---------------------------------------------------------------
' Re-points every Power Query connection to the current OneDrive directory
' (via QueriesAddress in Module2), then refreshes each connection
' synchronously. Returns only once every query has finished, so the calling
' Python code does not need a `time.sleep()` to wait for background queries.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the refresh; the original Application state is restored
' in the Cleanup block whether the Sub succeeded or raised an error.
'
' Side effect: each WorkbookConnection's BackgroundQuery flag is set to
' False and persists in the saved workbook. This is intentional.
Sub RefreshAll()

    Dim conn As WorkbookConnection
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Call QueriesAddress

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        On Error GoTo Cleanup
        conn.Refresh
    Next conn

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "RefreshAll error " & Err.Number & ": " & Err.Description, vbExclamation
    End If

End Sub


' --- SortAll ------------------------------------------------------------------
' Sorts the Pending_Offers ListObject ascending by Cx Offer, then by Account.
' Used after pending offers are loaded so the worksheet groups offers by
' account in the order the Python script attends them.
Sub SortAll()

    Dim prevCalc As Long
    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set POSheet = ThisWorkbook.Sheets(1)

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Clear

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Add2 Key:=Range("Pending_Offers[Cx Offer]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Add2 Key:=Range("Pending_Offers[Account]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With POSheet.ListObjects("Pending_Offers").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "SortAll error " & Err.Number & ": " & Err.Description, vbExclamation
    End If

End Sub


' --- Reorganize ---------------------------------------------------------------
' Final tidy-up sort once the offers have been attended: orders by Date
' (descending), then Account (ascending), then SKU (ascending). Called
' immediately before the workbook is saved and emailed.
Sub Reorganize()

    Dim prevCalc As Long
    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set POSheet = ThisWorkbook.Sheets(1)

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Clear

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Add2 Key:=Range("Pending_Offers[Date]"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Add2 Key:=Range("Pending_Offers[Account]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    POSheet.ListObjects("Pending_Offers").Sort.SortFields.Add2 Key:=Range("Pending_Offers[SKU]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With POSheet.ListObjects("Pending_Offers").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Reorganize error " & Err.Number & ": " & Err.Description, vbExclamation
    End If

End Sub
