Attribute VB_Name = "StringUtility"
'
' Module StringUtility
'
' Utility methods for working with strings
'

Option Explicit

Function Join(arrayToJoin() As String, Optional sep As String = ",", Optional stringDelim As String = "") As String
    On Error GoTo Catch
    CallStack.EnterRoutine "StringUtility.Join"
    
    Dim temp As String
    Dim i As Long
    temp = ""
    
    For i = LBound(arrayToJoin) To UBound(arrayToJoin) Step 1
        temp = temp & stringDelim & arrayToJoin(i) & stringDelim & sep
    Next i
    
    temp = Left(temp, Len(temp) - Len(sep))
    Join = temp
    
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
