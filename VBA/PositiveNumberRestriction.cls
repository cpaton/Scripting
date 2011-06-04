VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositiveNumberRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class PositiveNumberRestriction
'
' Used to restrict the data that can be inserted into a cell to a
' positive integer
'

Option Explicit

Public columnName As String
Public InputTooltip As String
Public AllowEmpty As Boolean
Public allowZero As Boolean

Public Sub ApplyToColumn(firstCellInColumn As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "PositiveNumberRestriction.SetupFromRetrieveValues"
    
    CellValidation.RestrictToPositiveInteger _
        Utility.ColumnRange(firstCellInColumn), _
        allowZero, _
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
    CallStack.EnterRoutine "PositiveNumberRestriction.IsValid"
    
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
    CallStack.EnterRoutine "PositiveNumberRestriction.InvalidReason"
    
    Dim cellValue As Date
    Dim errorMessage As String
    InvalidReason = ""

    If (IsNull(cell) Or IsEmpty(cell)) Then
        If (AllowEmpty) Then
            GoTo ExitBlock
        End If
        
        InvalidReason = columnName & " must be provided."
        GoTo ExitBlock
    End If
    
    If (allowZero) Then
        errorMessage = columnName & " must be a whole number greater than or equal to 0."
    Else
        errorMessage = columnName & " must be a whole number greater than 0."
    End If
    
    If (Not IsNumeric(cell.value)) Then
        InvalidReason = errorMessage
        GoTo ExitBlock
    End If
    
    If (allowZero) Then
        If (CLng(cell.value) = 0) Then
            InvalidReason = errorMessage
        End If
    Else
        If (CLng(cell.value) <= 0) Then
            InvalidReason = errorMessage
        End If
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
    CallStack.EnterRoutine "PositiveNumberRestriction.Class_Initialize"
    
    AllowEmpty = True
    columnName = ""
    InputTooltip = ""
    allowZero = False
    
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

