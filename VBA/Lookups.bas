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
Public Const BoxCodesRangeName As String = "BoxCodes"
Public Const IsoCodesRangeName As String = "IsoCodes"
Public Const OwnersRangeName As String = "Owners"
Public Const ReportingDateConversionTypesRangeName As String = "ConversionTypes"
Public Const DisclosureLevelsRangeName As String = "DisclosureLevels"
Public Const DenominationRangeName As String = "Denominations"

Private Enum LookupColumns
    LookupColumns_BankCodes = 1
    LookupColumns_BoxCodes = 3
    LookupColumns_IsoCodes = 5
    LookupColumns_Owners = 7
    LookupColumns_DisclosureLevels = 9
    LookupColumns_ConversionTypes = 11
    LookupColumns_Denominations = 13
End Enum

Public Sub RefreshLookupValues()
    On Error GoTo Catch
    CallStack.EnterRoutine "Lookups.RefreshLookupValues"
    
    Dim lookupValues As LookupDetails
    
    Set lookupValues = LookupsDA.GetLookupValues()
    
    RefreshLookup BankCodesRangeName, LookupColumns_BankCodes, lookupValues.BankCodes
    RefreshLookup BoxCodesRangeName, LookupColumns_BoxCodes, lookupValues.BoxCodes
    RefreshLookup IsoCodesRangeName, LookupColumns_IsoCodes, lookupValues.IsoCodes
    RefreshLookup OwnersRangeName, LookupColumns_Owners, lookupValues.Owners
    RefreshLookup ReportingDateConversionTypesRangeName, LookupColumns_ConversionTypes, lookupValues.DateConversionTypes
    RefreshLookup DisclosureLevelsRangeName, LookupColumns_DisclosureLevels, lookupValues.DisclosureLevels
    RefreshLookup DenominationRangeName, LookupColumns_Denominations, lookupValues.Denominations
    
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
