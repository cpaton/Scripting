Attribute VB_Name = "Lookups"
'
' Module Lookups
'
' Routines for refreshing the lookup values used to provide in cell drop downs
' of valid values that can be entered by the user.
'

Option Explicit

Public Const LookupSheetName As String = "Lookups"
Public Const BankCodesRangeName As String = "BankCodes"


Private Enum LookupColumns
    LookupColumns_BankCodes = 1
End Enum

Public Sub RefreshLookupValues()
    On Error GoTo Catch
    CallStack.EnterRoutine "Lookups.RefreshLookupValues"
    
    Dim db As DatabaseResult
    
    RefreshLookup BankCodesRangeName, LookupColumns_BankCodes, db
    
ExitBlock:
    On Error Resume Next
    CallStack.ExitRoutine
    Exit Sub
    
Catch:
    Select Case ErrorPolicy.HandleError(Err.Number, Err.Description, CallStack.CurrentRountine)
        Case ErrorPolicy.ResumeFromNextLine
            Resume Next
        Case ErrorPolicy.ExitCurrentRoutine
            Resume ExitBlock
        Case ErrorPolicy.StopExecution
            End
    End Select
End Sub

Private Sub RefreshLookup(rangeName As String, columnIndex As Long, rangeValues As DatabaseResult)
    On Error GoTo Catch
    CallStack.EnterRoutine "Lookups.RefreshLookup"
    
    Dim lookupWorksheet As Worksheet
    Dim rangeColumn As Range
    Dim firstCellInRange As Range
    Dim lookupRange As Range
    
    Utility.DeleteNamedRange rangeName

    Set lookupWorksheet = Utility.GetWorksheetByName(Lookups.LookupSheetName)
    Set rangeColumn = lookupWorksheet.Columns(columnIndex)
    Set firstCellInRange = lookupWorksheet.Cells(2, columnIndex)
    
    rangeColumn.Clear
    lookupWorksheet.Cells(1, columnIndex).value = rangeName
    firstCellInRange.CopyFromRecordset rangeValues.result
    
    Set lookupRange = lookupWorksheet.Range(firstCellInRange, firstCellInRange.End(xlDown))
    Utility.CreateNamedRange rangeName, lookupWorksheet, lookupRange
    
ExitBlock:
    On Error Resume Next
    CallStack.ExitRoutine
    Exit Sub
    
Catch:
    Select Case ErrorPolicy.HandleError(Err.Number, Err.Description, CallStack.CurrentRountine)
        Case ErrorPolicy.ResumeFromNextLine
            Resume Next
        Case ErrorPolicy.ExitCurrentRoutine
            Resume ExitBlock
        Case ErrorPolicy.StopExecution
            End
    End Select
End Sub
