Attribute VB_Name = "Formats"
'
' Module Formats
'
' Utility routines to format data that is shown to the user
' so that it is consistent throughout the workbook
'

Option Explicit

Public Const UserDateFormat As String = "dd-MMM-yyyy"
Public Const TextFormat As String = "@"

Public Function DateFormat(dateToFormat As Date) As String
    DateFormat = Format(dateToFormat, UserDateFormat)
End Function
