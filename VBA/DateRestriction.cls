VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class DateRestriction
'
' Used to restrict the data that can be inserted into a cell to a certain
' data range
'

Option Explicit

Public MustBeAfter As Date
Public MustBeOnOrBefore As Date
Public columnName As String
Public InputTooltip As String
Public AllowEmpty As Boolean

Public Sub ApplyToColumn(firstCellInColumn As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "DateRestriction.ApplyToColumn"
    
    CellValidation.RestrictToDateRange _
        Utility.ColumnRange(firstCellInColumn), _
        startDate:=MustBeAfter, _
        endDate:=MustBeOnOrBefore, _
        helpTitle:=columnName, _
        helpMessage:=InputTooltip
    
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
    CallStack.EnterRoutine "DateRestriction.IsValid"
    
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
    CallStack.EnterRoutine "DateRestriction.InvalidReason"
    
    Dim cellValue As Date
    InvalidReason = ""

    If (IsNull(cell) Or IsEmpty(cell)) Then
        If (AllowEmpty) Then
            GoTo ExitBlock
        End If
        
        InvalidReason = columnName & " must be provided."
        GoTo ExitBlock
    End If
    
    If (Not IsDate(cell.value)) Then
        InvalidReason = "The given " & columnName & " is not a valid date"
        GoTo ExitBlock
    End If
    
    cellValue = CDate(cell.value)
    If (Not cell.Validation.value) Then
        InvalidReason = "The given " & columnName & " '" & Formats.DateFormat(cellValue) & _
            "' is not in the allowed range " & Formats.DateFormat(MustBeAfter) & _
            " to " & Formats.DateFormat(MustBeOnOrBefore) & "."
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
    CallStack.EnterRoutine "DateRestriction.Class_Initialize"
    
    AllowEmpty = True
    MustBeAfter = Constants.MidasStartDate
    MustBeOnOrBefore = DateAdd("d", -1, Constants.MidasEndDate)
    InputTooltip = ""
    columnName = ""
    AllowEmpty = True
    
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
