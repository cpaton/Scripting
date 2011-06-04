VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumericRangeRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class NumericRangeRestriction
'
' Used to restrict the data that can be inserted into a cell to a
' number in a specified range
'

Option Explicit

Public columnName As String
Public InputTooltip As String
Public AllowEmpty As Boolean
Public MinimumAllowed As Long
Public MaximimAllowed As Long

Public Sub ApplyToColumn(firstCellInColumn As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "NumericRangeRestriction.SetupFromRetrieveValues"
    
    CellValidation.RestrictToIntegerRange _
        Utility.ColumnRange(firstCellInColumn), _
        MinimumAllowed, _
        MaximimAllowed, _
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
    CallStack.EnterRoutine "NumericRangeRestriction.IsValid"
    
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
    CallStack.EnterRoutine "NumericRangeRestriction.InvalidReason"
    
    Dim cellValue As Integer
    InvalidReason = ""

    If (IsNull(cell) Or IsEmpty(cell)) Then
        If (AllowEmpty) Then
            GoTo ExitBlock
        End If
        
        InvalidReason = columnName & " must be provided."
        GoTo ExitBlock
    End If
    
    If (Not IsNumeric(cell.value)) Then
        InvalidReason = columnName & " must be a number between " & MinimumAllowed & " and " & MaximimAllowed
        GoTo ExitBlock
    End If
    
    cellValue = CLng(cell.value)
    
    If (cellValue < MinimumAllowed Or MaximimAllowed < cellValue) Then
        InvalidReason = columnName & " must be a number between " & MinimumAllowed & " and " & MaximimAllowed
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
    CallStack.EnterRoutine "NumericRangeRestriction.Class_Initialize"
    
    AllowEmpty = True
    columnName = ""
    InputTooltip = ""
    MinimumAllowed = 0
    
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



