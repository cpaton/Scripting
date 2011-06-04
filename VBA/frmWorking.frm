VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorking 
   Caption         =   "{ Title }"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   OleObjectBlob   =   "frmWorking.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWorking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' frmWorking
'
' Non-modal form displayed to the user to provide feedback and progress when a long
' running action is in progress.  Only one of these forms should be open at once.  The
' opening and closing of this form is controlled in the module Progress.
'

Option Explicit

Private lastProgressUpdate As Date

Public Sub Initialise(title As String, task As String, initialProgress As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "frmWorking.ReportProgress"
    
    caption = title
    lblTask = task
    lblProgress = initialProgress
    lastProgressUpdate = Now()
    
    Repaint
    
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

Public Sub ReportProgress(currentProgress As String, Optional newTask As String = "")
    On Error GoTo Catch
    CallStack.EnterRoutine "frmWorking.ReportProgress"
    
    lastProgressUpdate = Now()
    
    If (newTask <> "") Then
        lblTask = newTask
    End If
    
    lblProgress = currentProgress
    
    Repaint
    DoEvents
    
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo Catch
    CallStack.EnterRoutine "frmWorking.UserForm_QueryClose"

    Dim oneMinuteAgo As Date

    If (CloseMode = VbQueryClose.vbFormControlMenu) Then
        '
        ' Don't allow the user to close the form is we are still receving progress updates
        '
        oneMinuteAgo = DateAdd("n", -1, Now())
        If (lastProgressUpdate > oneMinuteAgo) Then
            Cancel = True
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

Private Sub UserForm_Terminate()
    Progress.FormClosed
End Sub
