Attribute VB_Name = "Progress"
'
' Module Progress
'
' Controls display progress updates of long running tasks to users.  Progress is
' shown by frmWorking.  Low level routines have been coded to always report progress
' but this is only shown to the user if the higher level routine has decided to
' report progress by calling StartTask.  Once the task is complete the progress
' message can be hidden by calling TaskComplete
'
'

Option Explicit

Private g_workingForm As frmWorking

Public Sub StartTask(title As String, task As String, Optional initialProgress As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "Progress.ShowWorkingMessage"
    
    If (g_workingForm Is Nothing) Then
        Set g_workingForm = New frmWorking
    End If
    
    g_workingForm.Show
    g_workingForm.Initialise title, task, initialProgress
    
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

Public Sub Update(currentProgress As String, Optional newTask As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "Progress.Update"
    
    If (g_workingForm Is Nothing) Then
        GoTo ExitBlock
    End If
    
    g_workingForm.ReportProgress currentProgress, newTask
    
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

Public Sub TaskComplete()
    On Error GoTo Catch
    CallStack.EnterRoutine "Progress.TaskComplete"
    
    Progress.EnsureHidden
    
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

Public Sub EnsureHidden()
    On Error GoTo Catch
    CallStack.EnterRoutine "Progress.EnsureHidden"
    
    If (Not g_workingForm Is Nothing) Then
        g_workingForm.Hide
        Unload g_workingForm
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

Public Sub FormClosed()
    On Error GoTo Catch
    CallStack.EnterRoutine "Progress.FormClosed"
    
    Set g_workingForm = Nothing
    
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
