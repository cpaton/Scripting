Attribute VB_Name = "SourceControl"
'
' Module SourceControl
'
' Only used during development of the loader spreadsheet where the routine ExtractCode
' is run manually when code is ready to be checked into source control.  It extracts
' the source code from each of the modules into a file with the modules name into
' the same directory as the workbook.  These individual files should then be checked
' into source control.  Having individual files ensures history information is built
' up in TFS
'
' Before making changes to the spreadsheet the entire Excel Loader directory should be
' checked out.  When the changes are complete the ExtractCode macro should be run.  After
' that any files that have not changed should have their checkout undone so that they
' are not included in the changeset. This can be done automatically using the Team
' Foundation Power Tools command line tool tfpt which can be found in Source Control
' under the Tools\TfsPowerTools. Open a command prompt at the loader directory and type
'
' ..\..\..\Tools\TfsPowerTools\tfpt uu .
'

Option Explicit

Public Sub ExtractCode()
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.ExtractCode"
    
    Dim project As VBProject
    Dim codeModule As VBComponent
    Dim extractedModules As String
    
    ActiveWorkbook.Save
    CompileCode
    
    Set project = ActiveWorkbook.VBProject
    extractedModules = ""
    
    For Each codeModule In project.VBComponents
        If (codeModule.codeModule.CountOfLines > 0) Then
            extractedModules = extractedModules & ExportCodeModule(codeModule) & vbCrLf
        Else
            RemoveExportedCodeModule codeModule
        End If
    Next codeModule
    
    MsgBox "Code successfully extracted!" & vbCrLf & vbCrLf & extractedModules, vbInformation
    
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

Private Function ExportCodeModule(codeModule As VBComponent) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.ExportCodeModule"

    Dim fileSystem As New FileSystemObject
    Dim codeFilename As String
    Dim existingFile As File
    
    codeFilename = ExportedFilenameForModule(codeModule)
    
    If (fileSystem.FileExists(codeFilename)) Then
        Set existingFile = fileSystem.GetFile(codeFilename)
        If ((existingFile.Attributes And ReadOnly) = ReadOnly) Then
            MsgBox "The code file " & codeModule.name & " cannot be extracted as the file " & codeFilename & " is readonly.  Check the files out from source control and try again", vbCritical
            GoTo ExitBlock:
        End If
    End If
    
    codeModule.Export codeFilename
    ExportCodeModule = codeFilename
    
ExitBlock:
    On Error Resume Next
    Set existingFile = Nothing
    Set fileSystem = Nothing
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

Private Sub RemoveExportedCodeModule(codeModule As VBComponent)
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.RemoveExportedCodeModule"

    Dim fileSystem As New FileSystemObject
    Dim codeFilename As String
    Dim existingFile As File
    
    codeFilename = ExportedFilenameForModule(codeModule)
    
    If (fileSystem.FileExists(codeFilename)) Then
        fileSystem.DeleteFile codeFilename, True
    End If
    
ExitBlock:
    On Error Resume Next
    Set fileSystem = Nothing
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

Private Function ExportedFilenameForModule(codeModule As VBComponent) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.ExportedFilenameForModule"
    
    Dim filename As String
    Dim sheet As Worksheet
    
    If (codeModule.Type = vbext_ct_Document) Then
        '
        ' exporting a sheet or a workbook if it is a sheet add the sheet name to the file name
        '
        If (UCase(Left(codeModule.name, 5)) = "SHEET") Then
            For Each sheet In ActiveWorkbook.Sheets
                If (sheet.CodeName = codeModule.name) Then
                    filename = codeModule.name & "(" & sheet.name & ")"
                End If
            Next sheet
        Else
            filename = codeModule.name
        End If
    Else
        filename = codeModule.name
    End If

    ExportedFilenameForModule = ActiveWorkbook.Path & "\" & filename & "." & GetExtension(codeModule.Type)
    
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

Private Function GetExtension(moduleType As VBIDE.vbext_ComponentType) As String
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.GetExtension"
    
    Select Case moduleType
        Case vbext_ct_ClassModule
            GetExtension = "cls"
        Case vbext_ct_Document
            GetExtension = "cls"
        Case VBIDE.vbext_ComponentType.vbext_ct_MSForm
            GetExtension = "frm"
        Case Else
            GetExtension = "bas"
    End Select
    
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

Public Sub CompileCode()
    On Error GoTo Catch
    CallStack.EnterRoutine "SourceControl.CompileCode"
    
    Dim project As VBProject
    Dim toolbar As CommandBar
    Dim menus As CommandBar
    Dim menu As CommandBarControl
    Dim debugMenu As CommandBarPopup
    Dim menuItem As CommandBarControl
    Dim compileMenuItem As CommandBarButton
    
    Set project = ActiveWorkbook.VBProject
    
    For Each toolbar In project.VBE.CommandBars
        If (toolbar.name = "Menu Bar") Then
            Set menus = toolbar
            Exit For
        End If
    Next toolbar
    
    For Each menu In menus.Controls
        If (UCase(Replace(menu.Caption, "&", "")) = "DEBUG") Then
            Set debugMenu = menu
            Exit For
        End If
    Next
    
    For Each menuItem In debugMenu.Controls
        If (UCase(Replace(menuItem.Caption, "&", "")) = "COMPILE VBAPROJECT") Then
            Set compileMenuItem = menuItem
            Exit For
        End If
    Next menuItem
    
    If (compileMenuItem.Enabled) Then
        compileMenuItem.Execute
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
