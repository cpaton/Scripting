Attribute VB_Name = "Transaction"
'
' Module Transaction
'
' Provides a centralised place for code to start, commit and rollback database transactions
'

Option Explicit

Private g_activeTransaction As Boolean
Private g_transactionConnection As ADODB.connection

Public Sub Start()
    On Error GoTo Catch
    CallStack.EnterRoutine "Transaction.TransactionStarted"
    
    DataAccess.BeginTransaction
    
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

Public Sub TransactionStarted(databaseConnection As ADODB.connection)
    On Error GoTo Catch
    CallStack.EnterRoutine "Transaction.TransactionStarted"
    
    If (g_activeTransaction) Then
        ErrorPolicy.RaiseError ErrorCodes_AlreadyInTransaction, _
            "There is already a transaction active."
    End If
    
    g_activeTransaction = True
    Set g_transactionConnection = databaseConnection
    
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

Public Sub Commit()
    On Error GoTo Catch
    CallStack.EnterRoutine "Transaction.Commit"
    
    If (Not g_activeTransaction) Then
        ErrorPolicy.RaiseError ErrorCodes_NoActiveTransaction, _
            "There is no active transaction to commit."
    End If
    
    g_transactionConnection.CommitTrans
    g_activeTransaction = False
    Set g_transactionConnection = Nothing
    
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

Public Sub Rollback()
    On Error GoTo Catch
    CallStack.EnterRoutine "Transaction.Rollback"
    
    If (Not g_activeTransaction) Then
        ErrorPolicy.RaiseError ErrorCodes_NoActiveTransaction, _
            "There is no active transaction to rollback."
    End If
    
    g_transactionConnection.RollbackTrans
    g_activeTransaction = False
    Set g_transactionConnection = Nothing
    
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

Public Sub RollbackIfThereIsAnActiveTransaction()
    On Error GoTo Catch
    CallStack.EnterRoutine "Transaction.RollbackIfThereIsAnActiveTransaction"
    
    If (Not g_activeTransaction) Then
        GoTo ExitBlock
    End If
    
    Transaction.Rollback
    
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


