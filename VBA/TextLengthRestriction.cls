VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextLengthRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class TextLengthRestriction
'
' Used to restrict the data that can be inserted into a cell to a
' string that doesn't have too many characters
'

Option Explicit

Public MaximumLength As Long
Public AllowEmpty As Boolean
Public columnName As String
Public InputTooltip As String

Public Sub ApplyToColumn(firstCellInColumn As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "TextLengthRestriction.ApplyToColumn"
    
    CellValidation.RestrictTextLength _
        Utility.ColumnRange(firstCellInColumn), _
        35, _
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
    CallStack.EnterRoutine "TextLengthRestriction.IsValid"
    
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
    CallStack.EnterRoutine "TextLengthRestriction.InvalidReason"
    
    Dim cellValue As Date
    InvalidReason = ""

    If (IsNull(cell) Or IsEmpty(cell)) Then
        If (AllowEmpty) Then
            GoTo ExitBlock
        End If
        
        InvalidReason = columnName & " must be provided."
        GoTo ExitBlock
    End If
    
    If (Not cell.Validation.value) Then
        InvalidReason = columnName & " can not be longer than " & MaximumLength & " characters."
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
    CallStack.EnterRoutine "TextLengthRestriction.Class_Initialize"
    
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
