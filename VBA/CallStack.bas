Attribute VB_Name = "CallStack"
'
' Module CallStack
'
' VBA implementation of a call stack that keeps track of all functions
' or subs that have been called.  This information is used to provide
' better error messages and is analagous to the call stack found in
' .Net exceptions.
'
' The call stack is not maintained automatically it requires each sub
' and function to notify the call stack when it is entered or exited.
' These calls are made in the standard error handling blocks that are
' added to each method by convention
'

Option Explicit

Private g_stackTop As Long
Private g_callStack() As String
Private Const StackIncrement As Long = 10

Public Sub EnterRoutine(routineName As String)
    On Error Resume Next
    
    If (g_stackTop = UBound(g_callStack)) Then
        If (Err.Number = 9) Then
            '
            ' stack has not been initialised yet
            '
            Err.Clear
            g_stackTop = 0
            ReDim g_callStack(1 To StackIncrement)
        Else
            ReDim Preserve g_callStack(UBound(g_callStack) + StackIncrement)
        End If
    End If
    
    g_stackTop = g_stackTop + 1
    g_callStack(g_stackTop) = routineName
End Sub

Public Sub ExitRoutine()
    On Error Resume Next
    
    If (g_stackTop < UBound(g_callStack) And g_stackTop >= LBound(g_callStack)) Then
        g_callStack(g_stackTop) = ""
    End If
    
    g_stackTop = g_stackTop - 1
    If (g_stackTop < LBound(g_callStack)) Then
        g_stackTop = LBound(g_callStack) - 1
    End If
End Sub

Public Function CurrentRountine() As String
    CurrentRountine = g_callStack(g_stackTop)
End Function

Public Function GetCallStack(Optional itemSeperator As String = vbCrLf & vbTab) As String
    On Error Resume Next
    
    Dim currentEntry As String
    Dim index As Long
    Dim currentCallStack As String
    
    If (g_stackTop < LBound(g_callStack)) Then
        GetCallStack = ""
        Exit Function
    End If
    
    currentCallStack = ""
    For index = g_stackTop To LBound(g_callStack) Step -1
        currentEntry = g_callStack(index)
        currentCallStack = currentCallStack & currentEntry & itemSeperator
    Next index
    
    GetCallStack = Left(currentCallStack, Len(currentCallStack) - Len(itemSeperator))
End Function

Public Sub Clear()
    g_stackTop = LBound(g_callStack) - 1
End Sub

