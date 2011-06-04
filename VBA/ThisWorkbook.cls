VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Const ToolbarName As String = "Sample Toolbar"

Private Sub Workbook_Open()
    On Error Resume Next
    
    RemoveToolbar
    
    Dim toolbar As CommandBar
    Dim addInsertButton As CommandBarButton
    
    Set toolbar = Application.CommandBars.Add(ToolbarName)
    
    With toolbar
        .Visible = True
        .Position = msoBarTop
    End With
    
    AddButtonToToolbar toolbar, "Run Some Macro", "HelloWorld", 2950
    AddButtonToToolbar toolbar, "Extract Source Code", "SourceControl.ExtractCode", 620
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    RemoveToolbar
End Sub

Private Sub RemoveToolbar()
    On Error Resume Next
    Application.CommandBars(ToolbarName).Delete
End Sub

Private Function AddButtonToToolbar(toolbar As CommandBar, caption As String, subName As String, faceId As Integer) As CommandBarButton
    Dim button As CommandBarButton
    Dim fullAction
    
    fullAction = "'" + Me.FullName + "'!" + subName
    
    Set button = toolbar.Controls.Add(Type:=msoControlButton)
    With button
        .Visible = True
        .Style = msoButtonIconAndCaption
        .caption = caption
        .OnAction = fullAction
        .faceId = faceId
    End With
    Set AddButtonToToolbar = button
End Function





