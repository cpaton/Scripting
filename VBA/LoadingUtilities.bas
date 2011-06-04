Attribute VB_Name = "LoadingUtilities"
'
' Module LoadingUtilities
'
' Utility routines specific to loading Outputs data from a worksheet
'

Option Explicit

Public Sub IterateOverData(sheet As Worksheet, FirstDataRow As Long, _
        firstDataColumn As Long, _
        lastDataColumn As Long, _
        callback As Variant)
    On Error GoTo Catch
    CallStack.EnterRoutine "LoadingUtilities.IterateOverData"
    
    Dim currentRowIndex As Long
    Dim lastCell As Range
    Dim lastRowIndex As Long
    Dim columnIndex As Long
    
    lastRowIndex = 0
    
    For columnIndex = firstDataColumn To lastDataColumn
        Set lastCell = Utility.GetLastCellWithDataInColumn(sheet.Cells(FirstDataRow, columnIndex))
        If lastCell.Row > lastRowIndex Then
            lastRowIndex = lastCell.Row
        End If
    Next columnIndex
          
    For currentRowIndex = FirstDataRow To lastRowIndex
        If (Not IsRowEmpty(sheet, currentRowIndex, firstDataColumn, lastDataColumn)) Then
            If (VarType(callback) = vbObject) Then
                callback.ValidateRow sheet.Cells(currentRowIndex, 1)
            Else
                Application.Run callback, sheet.Cells(currentRowIndex, 1)
            End If
        End If
        If (currentRowIndex Mod 10 = 0 Or currentRowIndex = FirstDataRow) Then
            Progress.Update "Processing row " & currentRowIndex & " of " & lastRowIndex
        End If
    Next currentRowIndex

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

Private Function IsRowEmpty(sheet As Worksheet, RowIndex As Long, _
        firstDataColumn As Long, _
        lastDataColumn As Long) As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "LoadingUtilities.IsRowEmpty"
    
    Dim columnIndex As Long
    
    IsRowEmpty = True
     
    For columnIndex = firstDataColumn To lastDataColumn
        If Not IsEmpty(sheet.Cells(RowIndex, columnIndex)) Then
            IsRowEmpty = False
            Exit For
        End If
    Next columnIndex
    
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

Public Function GetNullableDate(cell As Range) As Variant
    On Error GoTo Catch
    CallStack.EnterRoutine "LoadingUtilities.GetNullableDate"
    
    Dim dateValue As Variant
    dateValue = Null
    
    If (IsDate(cell.value)) Then
        dateValue = CDate(cell.value)
    End If
    
    GetNullableDate = dateValue
    
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


