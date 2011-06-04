VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class LookupRestriction
'
' Used to restrict the data that can be inserted into a cell to
' only the values specified in a named range in the lookup sheet
'

Option Explicit

Public columnName As String
Public InputTooltip As String
Public AllowEmpty As Boolean
Public lookupRangeName As String

Public Sub ApplyToColumn(firstCellInColumn As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "ListRestriction.ApplyToColumn"
    
    CellValidation.RestrictToLookupRange _
        Utility.ColumnRange(firstCellInColumn), _
        lookupRangeName, _
        columnName, _
        columnName, _
        InputTooltip
    
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

Public Function IsValid(cell As Range) As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "ListRestriction.IsValid"
    
    IsValid = InvalidReason(cell) = ""
    
ExitBlock:
    On Error Resume Next
    CallStack.ExitRoutine
    Exit Function
    
Catch:
    Select Case ErrorPolicy.HandleError(Err.Number, Err.Description, CallStack.CurrentRountine)
        Case ErrorPolicy.ResumeFromNextLine
            Resume Next
        Case ErrorPolicy.ExitCurrentRoutine
            Resume ExitBlock
        Case ErrorPolicy.StopExecution
            End
    End Select
End Function

Public Function InvalidReason(cell As Range) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "ListRestriction.InvalidReason"
    
    Dim isValidValue As Boolean
    InvalidReason = ""

    If (IsNull(cell) Or IsEmpty(cell)) Then
        If (AllowEmpty) Then
            GoTo ExitBlock
        End If
        
        InvalidReason = columnName & " must be provided."
        GoTo ExitBlock
    End If
    
    isValidValue = cell.Validation.value
        
    If (Not isValidValue) Then
        InvalidReason = cell.value & " is not a valid " & columnName & ". Check the in cell dropdown for valid values."
    End If
    
ExitBlock:
    On Error Resume Next
    CallStack.ExitRoutine
    Exit Function
    
Catch:
    Select Case ErrorPolicy.HandleError(Err.Number, Err.Description, CallStack.CurrentRountine)
        Case ErrorPolicy.ResumeFromNextLine
            Resume Next
        Case ErrorPolicy.ExitCurrentRoutine
            Resume ExitBlock
        Case ErrorPolicy.StopExecution
            End
    End Select
End Function

Private Sub Class_Initialize()
    On Error GoTo Catch
    CallStack.EnterRoutine "ListRestriction.Class_Initialize"
    
    AllowEmpty = True
    columnName = ""
    InputTooltip = ""
    
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






