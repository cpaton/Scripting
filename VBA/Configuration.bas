Attribute VB_Name = "Configuration"
'
' Module Configuration
'
' Module used to store persistent configuration information in the form
' of name value pairs.  This module provides configuration capabilities
' similar to app.config in .Net.
'
' All configuration values are stored in the hidden configuration worksheet.
'
' The parameter values can be optionally "encrypted" so that they are not
' stored in plain text.  As this is VBA the level of encryption is low
' and should not be used to store any sensitive information
'

Option Explicit

Public Const ConfigurationSheetName As String = "Configuration"
Public Const DataSourceParameterName As String = "Data Source"
Public Const OptimizerParameterName As String = "Optimizer"
Public Const LastLoadTimeParameterNameSuffix As String = "Last Load Time"
Public Const OutputIdParameterName As String = "Output Definition Id"
Public Const CoreDataSetIdParameterName As String = "Core Data Definition Id"
Public Const UserRoleParameterName As String = "User Role"
Public Const UserRolePasswordParameterName As String = "Role Password"

Public Function GetParameterValue(parameterName As String, Optional encrypted As Boolean = False) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "Configuration.GetParameterValue"
    
    Dim parameterCell As Range
    Dim parameterValue As String
    
    Set parameterCell = GetParameterCell(parameterName)
    
    If (parameterCell Is Nothing) Then
        Err.Raise vbObjectError + 1, Description:="Unknown parameter - " & parameterName
    End If
    
    parameterValue = parameterCell.offset(0, 1).value
    If (encrypted) Then
        parameterValue = Base64.Base64DecodeString(parameterValue)
    End If
    GetParameterValue = parameterValue
    
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

Public Sub SetParameterValue(parameterName As String, parameterValue As Variant, Optional encrypt As Boolean = False)
    On Error GoTo Catch
    CallStack.EnterRoutine "Configuration.SetParameterValue"

    Dim parameterCell As Range
    Dim configurationSheet As Worksheet
    
    Set parameterCell = GetParameterCell(parameterName)
    
    If (parameterCell Is Nothing) Then
        Set configurationSheet = Utility.GetWorksheetByName(ConfigurationSheetName)
        Set parameterCell = Utility.GetNextEmptyCellInColumn(configurationSheet.Cells(2, 1))
    End If
    
    If (encrypt) Then
        parameterValue = Base64.Base64EncodeString(parameterValue)
    End If
    parameterCell.value = parameterName
    parameterCell.offset(0, 1).value = parameterValue
    
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

Private Function GetParameterCell(parameterName As String) As Variant
    On Error GoTo Catch
    CallStack.EnterRoutine "Configuration.GetParameterCell"

    Dim configurationSheet As Worksheet
    Dim nameRange As Range
    Dim parameterCell As Range

    Set configurationSheet = Utility.GetWorksheetByName(ConfigurationSheetName)
    Set nameRange = configurationSheet.Columns("A:A")
    Set GetParameterCell = nameRange.Find(What:=parameterName, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
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
