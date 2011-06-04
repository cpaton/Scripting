Attribute VB_Name = "ErrorPolicy"
'
' Module ErrorPolicy
'
' Routines implementing the standard error policy for the entire workbook.
' When an error occurs the HandleError routine is called which displays a
' standard formatted message and ensures that any open transactions are
' rolled back.  The message includes a call stack detailing the trail of
' subs and functions that were called before the error was thrown
'
' By convention every sub and function must have the following code added
'
' Directly after the routine declaration
'   On Error GoTo Catch
'   CallStack.EnterRoutine "<Routine Name>"
'
' e.g.
'   On Error GoTo Catch
'   CallStack.EnterRoutine "CellValidation.RestrictToDateRange"
'
'
' At the end of every routine
'ExitBlock:
'    On Error Resume Next
'    CallStack.ExitRoutine
'    Exit Sub
'
'Catch:
'    Select Case ErrorPolicy.HandleError(Err.Number, Err.Description, CallStack.CurrentRountine)
'        Case ErrorPolicy.ResumeFromNextLine
'            Resume Next
'        Case ErrorPolicy.ExitCurrentRoutine
'            Resume ExitBlock
'        Case ErrorPolicy.StopExecution
'            End
'    End Select
'
'
' In the current implementation every error results in the StopExecution code
' being returned which halts the application
'

Option Explicit

Public Const ResumeFromNextLine As Long = 1
Public Const ExitCurrentRoutine As Long = 2
Public Const StopExecution As Long = 3
Public CurrentRoutine As String

Public Enum ErrorCodes
    ErrorCodes_AlreadyInTransaction = vbObjectError + 1
    ErrorCodes_NoActiveTransaction = vbObjectError + 2
    ErrorCodes_DoesNotSupportPopulatingMissingData = vbObjectError + 3
    ErrorCodes_DoesNotSupportGettingLastUpdateTime = vbObjectError + 4
    ErrorCodes_MessageMissing = vbObjectError + 5
    ErrorCodes_NotImplemented = vbObjectError + 6
End Enum

Public Function HandleError(errorCode As Long, errorDescription As String, location As String) As Long
    On Error Resume Next
    
    Dim currentCallStack As String
    
    currentCallStack = CallStack.GetCallStack()
    CallStack.Clear
    
    Transaction.RollbackIfThereIsAnActiveTransaction
    Progress.EnsureHidden
    
    MsgBox "Unhandled error.  Please contact MIDAS support." & _
        vbCrLf & vbCrLf & vbCrLf & _
        "Location:" & vbTab & location & vbCrLf & _
        "Message:" & vbTab & errorDescription & vbCrLf & _
        "Code:" & vbTab & errorCode & vbCrLf & _
        "Call Stack:" & vbCrLf & vbTab & currentCallStack, _
        Buttons:=vbCritical, _
        title:=Constants.GenericMessageBoxTitle
        
    Application.Cursor = xlDefault
    HandleError = ErrorPolicy.StopExecution
End Function

Public Function RaiseError(errorCode As ErrorCodes, message As String)
    Err.Raise errorCode, "MIDAS Outputs Loader", message
End Function
