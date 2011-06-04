Attribute VB_Name = "CellValidation"
'
' Module CellValidation
'
' Routines for applying excel validation (Data -> Validation) to columns within
' a worksheet that is being used to load data.  These routines are used by
' the "Restriction" class modules to restrict what the user can type in.  This
' provides the first layer of data validation to prevent users from entering
' data that is obviously wrong e.g. writing a string in a cell that expects a date.
'
' The validation routines have optional parameters that can be used to show a tooltip
' when a user selects a cell in the column to explain the meaning of the data
'

Option Explicit

Public Sub RestrictToDateRange(cellRangeToValidate As Range, Optional startDate As Date = Constants.MidasStartDate, _
        Optional endDate As Date = Constants.MidasEndDate, Optional helpTitle As String = "", _
        Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictToDateRange"
    
    Dim dayAfterLastAllowedDate As Date
    Dim errorMessage As String
    
    dayAfterLastAllowedDate = DateAdd("d", 1, endDate)
    errorMessage = "Date must be after " & Format(startDate, "dd-MMM-yyyy") & " and before " & Format(dayAfterLastAllowedDate, "dd-MMM-yyyy")
    
    With cellRangeToValidate.Validation
        .delete
        .Add xlValidateDate, xlValidAlertStop, xlBetween, startDate, endDate
        .IgnoreBlank = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .ErrorTitle = "Date not within range"
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With
    
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

Public Sub RestrictTextLength(cellRangeToValidate As Range, maxLength As Integer, Optional helpTitle As String = "", _
        Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictTextLength"

    Dim errorMessage As String
    errorMessage = "Must be less than " & maxLength + 1 & " characters."
    
    With cellRangeToValidate.Validation
        .delete
        .Add xlValidateTextLength, xlValidAlertStop, xlBetween, 1, maxLength
        .IgnoreBlank = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .ErrorTitle = "Text Too Long"
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With
    
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

Public Sub RestrictToPositiveInteger(cellRangeToValidate As Range, allowZero As Boolean, _
        Optional helpTitle As String = "", Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictToPositiveInteger"

    Dim errorMessage As String
    Dim lowestAcceptableNumer As Long
    
    If (allowZero) Then
        lowestAcceptableNumer = 0
        errorMessage = "Must be a whole number greater than or equal to 0."
    Else
        lowestAcceptableNumer = 1
        errorMessage = "Must be a whole number greater than 0."
    End If
    
    With cellRangeToValidate.Validation
        .delete
        .Add xlValidateWholeNumber, xlValidAlertStop, xlGreater, lowestAcceptableNumer - 1
        .IgnoreBlank = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .ErrorTitle = "Not a Positive Whole Number"
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With
    
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

Public Sub RestrictToIntegerRange(cellRangeToValidate As Range, minimum As Long, maximum As Long, _
        Optional helpTitle As String = "", Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictToIntegerRange"

    Dim errorMessage As String
    errorMessage = "Must be between " & minimum & " and " & maximum
    
    With cellRangeToValidate.Validation
        .delete
        .Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, minimum, maximum
        .IgnoreBlank = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .ErrorTitle = "Out of Range"
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With
    
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

Public Sub RestrictToYesNo(cellRangeToValidate As Range, Optional helpTitle As String = "", _
        Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictToYesNo"
    
    With cellRangeToValidate.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes, No"
        .IgnoreBlank = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .ErrorTitle = "Yes/No"
        .errorMessage = "Must enter Yes or No or leave blank."
        .ShowInput = True
        .ShowError = True
    End With
    
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

Public Sub RestrictToLookupRange(cellRangeToValidate As Range, lookupRangeName As String, _
        dataType As String, Optional helpTitle As String = "", Optional helpMessage As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "CellValidation.RestrictToLookupRange"
    
    With cellRangeToValidate.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & lookupRangeName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = helpTitle
        .inputMessage = helpMessage
        .errorMessage = "The entered value is not a valid " & dataType & ". Check the in cell dropdown for valid values."
        .ShowInput = True
        .ShowError = True
    End With
    
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

