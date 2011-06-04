Attribute VB_Name = "Validation"
'
' Module Validation
'
' Implements routines that are used to work with validation messages that
' are returned from running the Oracle ValidateAndLoad procedures.
'
' The validation messages are stored in the hidden worksheet named Validation
' which has a named range called Validation.  Using a named range makes it
' easier to create forms which display the details
'

Option Explicit

Public Const ValidationSheetName As String = "Validation"
Public Const ValidationNamedRange As String = "Validation"

Private Const FirstValidationRow As Long = 2
Private Const RowNumberColumn As String = "ROW_NUMBER"
Private Const ColumnNameColumn As String = "COLUMN_NAME"
Private Const MessageColumn As String = "MESSAGE"
Private Const MessageTypeColumn As String = "MESSAGE_TYPE"
Private Const MessageCodeColumn As String = "MESSAGE_CODE"
Private g_sheetValidated As Worksheet
Private g_columnNameMap As Dictionary

Public Enum ValidationMessageColumns
    ValidationMessageColumns_First = 1
    ValidationMessageColumns_RowNumber = 1
    ValidationMessageColumns_ColumnName = 2
    ValidationMessageColumns_MessageType = 3
    ValidationMessageColumns_MessageCode = 4
    ValidationMessageColumns_Message = 5
    ValidationMessageColumns_Column = 6
    ValidationMessageColumns_Sheet = 7
    ValidationMessageColumns_Last = 7
End Enum

Public Sub OutputValidationMessagesFromRecordset(validationMessages As DatabaseResult)
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.OutputValidationMessagesFromRecordset"
    
    Dim validationSheet As Worksheet
    Dim currentRowIndex As Long
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    
    Utility.ClearDownSheet validationSheet, FirstValidationRow
    
    currentRowIndex = FirstValidationRow
    
    If (Not validationMessages.IsClosed And Not validationMessages.IsEmpty) Then
        With validationMessages.result
            .MoveFirst
            While (Not .EOF)
                validationSheet.Cells(currentRowIndex, ValidationMessageColumns_RowNumber).value = .Fields(RowNumberColumn).value
                validationSheet.Cells(currentRowIndex, ValidationMessageColumns_ColumnName).value = .Fields(ColumnNameColumn).value
                validationSheet.Cells(currentRowIndex, ValidationMessageColumns_MessageType).value = .Fields(MessageTypeColumn).value
                validationSheet.Cells(currentRowIndex, ValidationMessageColumns_MessageCode).value = .Fields(MessageCodeColumn).value
                validationSheet.Cells(currentRowIndex, ValidationMessageColumns_Message).value = .Fields(MessageColumn).value
            
                .MoveNext
                currentRowIndex = currentRowIndex + 1
            Wend
        End With
    End If
    
    RecreateNamedRange
    
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

Public Function ContainsMessages() As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.ContainsMessages"
    
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    
    ContainsMessages = Not IsEmpty(validationSheet.Cells(FirstValidationRow, ValidationMessageColumns_Message))
    
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

Public Sub UpdateLocations(sheetValidated As Worksheet, columnNameMap As Dictionary)
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.UpdateLocations"
    
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
        
    Set g_sheetValidated = sheetValidated
    Set g_columnNameMap = columnNameMap
    
    LoadingUtilities.IterateOverData _
        validationSheet, _
        FirstValidationRow, _
        ValidationMessageColumns_First, _
        ValidationMessageColumns_Last, _
        "Validation.FillInLocationDetails"
    
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

Public Sub FillInLocationDetails(firstColumnOfMessage As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.FillInLocationDetails"
    
    Dim columnName As String
    
    firstColumnOfMessage.offset(0, ValidationMessageColumns_Sheet - 1).value = g_sheetValidated.name
    
    If (Not IsEmpty(firstColumnOfMessage.offset(0, ValidationMessageColumns_ColumnName - 1))) Then
        columnName = firstColumnOfMessage.offset(0, ValidationMessageColumns_ColumnName - 1).value
        If (g_columnNameMap.Exists(UCase(columnName))) Then
            firstColumnOfMessage.offset(0, ValidationMessageColumns_Column - 1).value = g_columnNameMap(columnName).index
        End If
    End If
    
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

Public Function ContainsOnlyWarnings() As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.ContainsOnlyWarnings"
    
    Dim currentRow As Long
    Dim messageCount As Long
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    
    If (Not Validation.ContainsMessages()) Then
        ContainsOnlyWarnings = False
        GoTo ExitBlock
    End If
    
    currentRow = FirstValidationRow
    While (Not IsEmpty(validationSheet.Cells(currentRow, ValidationMessageColumns_Message)))
        If (UCase(validationSheet.Cells(currentRow, ValidationMessageColumns_MessageType).value) <> "WARNING") Then
            ContainsOnlyWarnings = False
            GoTo ExitBlock
        End If
    
        currentRow = currentRow + 1
    Wend
    
    ContainsOnlyWarnings = True
    
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

Public Function ContainsMessageWithCode(messageCode As String) As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.ContainsMessageWithCode"
    
    Dim currentRow As Long
    Dim messageCount As Long
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    
    If (Not Validation.ContainsMessages() Or messageCode = "") Then
        ContainsMessageWithCode = False
        GoTo ExitBlock
    End If
    
    currentRow = FirstValidationRow
    While (Not IsEmpty(validationSheet.Cells(currentRow, ValidationMessageColumns_Message)))
        If (UCase(validationSheet.Cells(currentRow, ValidationMessageColumns_MessageCode).value) = UCase(messageCode)) Then
            ContainsMessageWithCode = True
            GoTo ExitBlock
        End If
    
        currentRow = currentRow + 1
    Wend
    
    ContainsMessageWithCode = False
    
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

Public Sub ClearMessagesWithCode(messageCode)
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.ClearMessagesWithCode"
    
    Dim currentRow As Long
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    
    If (Not Validation.ContainsMessages()) Then
        GoTo ExitBlock
    End If
    
    currentRow = FirstValidationRow
    While (Not IsEmpty(validationSheet.Cells(currentRow, ValidationMessageColumns_Message)))
        If (UCase(validationSheet.Cells(currentRow, ValidationMessageColumns_MessageCode).value) = UCase(messageCode)) Then
            validationSheet.Rows(currentRow).delete xlShiftUp
        Else
            currentRow = currentRow + 1
        End If
    Wend
    
    RecreateNamedRange
    
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

Private Sub RecreateNamedRange()
    On Error GoTo Catch
    CallStack.EnterRoutine "Validation.RecreateNamedRange"
    
    Dim validationSheet As Worksheet
    Set validationSheet = Utility.GetWorksheetByName(ValidationSheetName)
    Dim validationStartCell As Range
    Dim lastMessageCell As Range
    Dim validationMessagesRange As Range
    
    Utility.DeleteNamedRange ValidationNamedRange
    
    If (ContainsMessages) Then
        Set validationStartCell = validationSheet.Cells(FirstValidationRow, ValidationMessageColumns_First)
        Set lastMessageCell = Utility.GetLastCellWithDataInColumn( _
            validationSheet.Cells(FirstValidationRow, ValidationMessageColumns_Message))
        Set validationMessagesRange = validationSheet.Range( _
            validationStartCell, _
            validationSheet.Cells(lastMessageCell.Row, ValidationMessageColumns_Last))
        Utility.CreateNamedRange ValidationNamedRange, validationSheet, validationMessagesRange
    End If
    
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


