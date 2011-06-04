Attribute VB_Name = "Utility"
'
' Module Utility
'
' Routines that make working with the Excel object model easier
'

Option Explicit

Public Function GetWorksheetByName(worksheetName As String) As Worksheet
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.ExtractCode"
    
    Dim sheet As Worksheet
    
    For Each sheet In ActiveWorkbook.Sheets
        If (sheet.name = worksheetName) Then
            Set GetWorksheetByName = sheet
            GoTo ExitBlock
        End If
    Next sheet
    
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

Public Sub CreateNamedRange(rangeName As String, forSheet As Worksheet, forCellRange As Range)
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.CreateNamedRange"
    
    Dim rangeAddress As String
    rangeAddress = "=" & forSheet.name & "!" & forCellRange.Address
    ActiveWorkbook.Names.Add name:=rangeName, RefersTo:=rangeAddress
    
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

Public Sub DeleteNamedRange(rangeName As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.DeleteNamedRange"
    
    Dim rangeToDelete As Range
    On Error Resume Next
    Set rangeToDelete = ActiveWorkbook.Names(rangeName)
    rangeToDelete.Delete
    
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

Public Function ColumnRange(firstRow As Range) As Range
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.ColumnRange"
    
    Set ColumnRange = firstRow.Worksheet.Range(firstRow, firstRow.Worksheet.Cells(65536, firstRow.Column))
    
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

Public Sub ClearDownSheet(sheet As Worksheet, fromRow As Long)
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.ClearDownSheet"
    
    Dim dataStartCell As Range
    Dim dataRange As Range
    
    Set dataStartCell = sheet.Cells(fromRow, 1)
    Set dataRange = sheet.Range(dataStartCell, sheet.Cells(65536, 255))
    dataRange.ClearContents
    dataRange.ClearComments
    dataRange.Validation.Delete
    dataRange.ClearFormats
    
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

Public Sub HideAllWorksheetsExcept(worksheetName As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.HideAllWorksheetsExcept"
    
    Dim sheet As Worksheet
    
    Set sheet = Utility.GetWorksheetByName(worksheetName)
    sheet.Visible = xlSheetVisible
    
    For Each sheet In ActiveWorkbook.Sheets
        If (sheet.name <> worksheetName And sheet.Visible = xlSheetVisible) Then
            sheet.Visible = xlSheetHidden
        End If
    Next sheet
    
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

Public Sub HideAllWorksheetsExceptThese(worksheetNames() As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.HideAllWorksheetsExceptThese"
    
    Dim sheet As Worksheet
    Dim index As Long
    
    For index = LBound(worksheetNames) To UBound(worksheetNames)
        Set sheet = Utility.GetWorksheetByName(worksheetNames(index))
        sheet.Visible = xlSheetVisible
    Next index
    
    For Each sheet In ActiveWorkbook.Sheets
        If (Not Utility.Contains(worksheetNames, sheet.name) And sheet.Visible = xlSheetVisible) Then
            sheet.Visible = xlSheetHidden
        End If
    Next sheet
    
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

Public Function GetNextEmptyCellInColumn(startCell As Range) As Range
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.GetNextEmptyCellInColumn"
    
    If (IsEmpty(startCell)) Then
        Set GetNextEmptyCellInColumn = startCell
    Else
        Set GetNextEmptyCellInColumn = startCell.End(xlDown).offset(1, 0)
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

Public Function GetLastCellWithDataInColumn(startCell As Range) As Range
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.GetLastCellWithDataInColumn"
    
    Dim lastCellWithData As Range
    
    Set lastCellWithData = startCell.Worksheet.Cells(Constants.RowLimit, startCell.Column).End(xlUp)
    
    If (lastCellWithData.Row < startCell.Row) Then
        Set GetLastCellWithDataInColumn = startCell
    Else
        Set GetLastCellWithDataInColumn = lastCellWithData
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

Public Function YesNoToBool(valueToConvert As String, Optional valueIfEmpty As Boolean = False) As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.YesNoToBool"

    If (IsNull(valueToConvert) Or IsEmpty(valueToConvert)) Then
        YesNoToBool = valueIfEmpty
        GoTo ExitBlock
    End If
    
    valueToConvert = UCase(valueToConvert)
    If (valueToConvert = "YES" Or valueToConvert = "Y" Or valueToConvert = "TRUE" Or valueToConvert = "1") Then
        YesNoToBool = True
    Else
        YesNoToBool = False
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

Public Function GetTimestamp() As Date
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.GetTimestamp"
    
    GetTimestamp = DataAccess.GetTimeOnDatabaseServer()

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

Public Function VariantToStringArray(toConvert As Variant) As String()
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.VariantToStringArray"

    Dim converted() As String
    Dim index As Long
    
    If (UBound(toConvert) < 0) Then
        ReDim converted(0)
        VariantToStringArray = converted
        GoTo ExitBlock
    End If
    
    ReDim converted(UBound(toConvert))
    
    For index = LBound(toConvert) To UBound(toConvert) Step 1
        converted(index) = toConvert(index)
    Next index
    
    VariantToStringArray = converted
    
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

Public Function VariantToLongArray(toConvert As Variant) As Long()
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.VariantToLongArray"

    Dim converted() As Long
    Dim index As Long
    
    If (UBound(toConvert) < 0) Then
        ReDim converted(0)
        VariantToLongArray = converted
        GoTo ExitBlock
    End If
    
    ReDim converted(UBound(toConvert))
    
    For index = LBound(toConvert) To UBound(toConvert) Step 1
        converted(index) = CLng(toConvert(index))
    Next index
    
    VariantToLongArray = converted
    
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

Public Function Contains(arrayToCheck() As String, lookFor As String, Optional caseSensitive As Boolean = False) As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "Utility.Contains"

    Dim index As Long
    
    For index = LBound(arrayToCheck) To UBound(arrayToCheck)
        If (UCase(arrayToCheck(index)) = UCase(lookFor)) Then
            Contains = True
            GoTo ExitBlock
        End If
    Next index
    
    Contains = False
    
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
