VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Class DatabaseResult
'
' Wrapper around a recordset returned from Oracle that ensures
' that the recordset is closed.  Calling code can call Dispose when
' it has finished interogating the data
'

Option Explicit

Private command As ADODB.command
Private results As ADODB.Recordset
Private disposed As Boolean

Public Property Get result() As ADODB.Recordset
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.result"
    
    If (disposed) Then
        Err.Raise vbObjectError + 2, "Database results already disposed and cannot be accessed"
    End If
    Set result = results
    
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
End Property

Public Sub SetDatabaseResponse(dbCommand As ADODB.command, dbResults As ADODB.Recordset)
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.SetDatabaseResponse"

    Set command = dbCommand
    Set results = dbResults
    
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

Public Sub Dispose()
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.Dispose"

    If (Not disposed) Then
        If (Not IsClosed) Then
            results.Close
        End If
        Set results = Nothing
        Set command = Nothing
    End If
    disposed = True
    
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

Public Function IsClosed() As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.IsClosed"
    
    If (results Is Nothing) Then
        IsClosed = True
    Else
        IsClosed = results.State = adStateClosed
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

Public Function IsEmpty() As Boolean
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.IsEmpty"
    
    If (results Is Nothing) Then
        IsEmpty = True
    Else
        IsEmpty = (results.BOF And results.EOF)
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
    CallStack.EnterRoutine "DatabaseResult.Class_Initialize"

    disposed = False
    
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

Private Sub Class_Terminate()
    On Error GoTo Catch
    CallStack.EnterRoutine "DatabaseResult.Class_Terminate"

    If (Not disposed) Then
        Dispose
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
