Attribute VB_Name = "DataAccess"
'
' Module DataAccess
'
' Low level data access routines and helper methods. These routines should
' be called by the Data Access (DA) classes and not directly by other
' modules
'
' This module manages a single connection to the Oralce database that is opened
' on first demand and closed when the workbook closes.  Connection to the
' Oracle back end is made using the Oracle Proivder for OLE DB using integrated
' security.
'

Option Explicit

Private g_Connection As ADODB.connection

Public Sub ResetConnection()
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.ResetConnection"
    
    If (Not g_Connection Is Nothing) Then
        g_Connection.Close
    End If
        
    Set g_Connection = Nothing
    
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

Public Sub BeginTransaction()
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.BeginTransaction"

    Dim conn As ADODB.connection
    
    Set conn = GetConnection()
    conn.BeginTrans
    Transaction.TransactionStarted conn
    
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

Public Function GetTimeOnDatabaseServer() As Date
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.GetTimeOnDatabaseServer"
    
    Dim sysdateCommand As ADODB.command
    Dim sysdateResult As ADODB.Recordset
    Dim sysdate As Date
    
    Set sysdateCommand = CreateCommand("select sysdate from dual")
    sysdateCommand.CommandType = adCmdText
    Set sysdateResult = sysdateCommand.Execute()
    sysdate = CDate(sysdateResult.Fields("SYSDATE").value)
    
    GetTimeOnDatabaseServer = sysdate
    
ExitBlock:
    On Error Resume Next
    Set sysdateResult = Nothing
    Set sysdateCommand = Nothing
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

Public Function OpenRecordsetForTable(tableName As String) As ADODB.Recordset
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.OpenRecordsetForTable"
    
    Dim conn As ADODB.connection
    Dim rst As New ADODB.Recordset
    
    Set conn = GetConnection()
    rst.Open tableName, conn, adOpenStatic, adLockOptimistic, adCmdTable
    
    Set OpenRecordsetForTable = rst
    
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

Public Function CreateCommand(commandText As String) As ADODB.command
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.CreateCommand"
    
    Dim cmd As New ADODB.command
    Set cmd.ActiveConnection = GetConnection
    cmd.commandText = commandText
    cmd.CommandType = adCmdStoredProc
    Set CreateCommand = cmd
    
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

Private Function GetConnection() As ADODB.connection
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.GetConnection"

    Dim optimizer As String
    Dim userRole As String
    Dim rolePassword As String

    If (Not g_Connection Is Nothing) Then
        Set GetConnection = g_Connection
        GoTo ExitBlock
    End If
    
    Set g_Connection = New ADODB.connection
    g_Connection.ConnectionString = ConnectionString()
    g_Connection.Open
    
    'set the optimizer and role
    optimizer = Configuration.GetParameterValue(Configuration.OptimizerParameterName)
    userRole = Configuration.GetParameterValue(Configuration.UserRoleParameterName)
    rolePassword = Configuration.GetParameterValue(Configuration.UserRolePasswordParameterName, True)
    
    DataAccess.ExecuteNonQuery g_Connection, "ALTER SESSION SET optimizer_features_enable = '" & optimizer & "'"
    DataAccess.SetRole g_Connection, userRole, rolePassword
    
    Set GetConnection = g_Connection
    
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

Private Function ConnectionString() As String
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.ConnectionString"

    Dim dataSource As String
    dataSource = Configuration.GetParameterValue(Configuration.DataSourceParameterName)
    ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & dataSource & ";OSAuthent=1;PLSQLRSet=1"
    
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

Public Sub AddRecordCountParameter(toCommand As ADODB.command)
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.AddRecordCountParameter"

    Dim recordCountParameter As ADODB.parameter

    Set recordCountParameter = toCommand.CreateParameter("po_record_count", adInteger, adParamOutput)
    toCommand.Parameters.Append recordCountParameter
    
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

Public Sub SetRole(connection As ADODB.connection, roleName As String, rolePassword As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.SetRole"
    
    ExecuteNonQuery connection, "SET ROLE " & roleName & " IDENTIFIED BY " & rolePassword
    
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

Public Sub ExecuteNonQuery(connection As ADODB.connection, commandText As String)
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.ExecuteNonQuery"
    
    Dim command As New ADODB.command
    command.commandText = commandText
    command.CommandType = adCmdText
    command.ActiveConnection = connection
    
    command.Execute
    
ExitBlock:
    On Error Resume Next
    Set command = Nothing
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

Public Sub AddInputParameter(toCommand As ADODB.command, name As String, dataType As DataTypeEnum, value As Variant, Optional size As Variant = Null)
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.AddInputParameter"
    
    Dim parameter As ADODB.parameter
    
    Set parameter = toCommand.CreateParameter(name, dataType, adParamInput, value:=value)
    
    If (Not IsNull(size)) Then
        parameter.size = size
    End If
    toCommand.Parameters.Append parameter
    
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

Public Function AddOutputParameter(toCommand As ADODB.command, name As String, dataType As DataTypeEnum, Optional size As Variant = Null) As ADODB.parameter
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.AddOutputParameter"
    
    Dim parameter As ADODB.parameter
    
    Set parameter = toCommand.CreateParameter(name, dataType, adParamOutput)
    
    If (Not IsNull(size)) Then
        parameter.size = size
    End If
    toCommand.Parameters.Append parameter
    
    Set AddOutputParameter = parameter
    
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

Public Sub AddBooleanInputParameter(toCommand As ADODB.command, name As String, value As Boolean)
    AddInputParameter toCommand, name, adVarChar, BooleanParameterValue(value), 1
End Sub

Private Function BooleanParameterValue(bool As Boolean) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "DataAccess.BooleanParameterValue"
    
    If (bool) Then
        BooleanParameterValue = "Y"
    Else
        BooleanParameterValue = "N"
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


