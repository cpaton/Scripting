Attribute VB_Name = "Constants"
'
' Module Constants
'
' Central location for all constant values that do not fit neatly
' into other modules.
'

Option Explicit

Public Const MidasStartDate As Date = "01-Jan-1900"
Public Const MidasEndDate As Date = "31-Dec-2025"
Public Const RowLimit As Long = 65536
Public Const ColumnLimit As Long = 256
Public Const GenericMessageBoxTitle = "MIDAS Outputs Loader"

Public Const MessageCode_MissingCoreDataSeries As String = "MISSING CORE DATA SERIES"
Public Const MessageCode_MissingExternalCodes As String = "MISSING EXTERNAL CODES"

Public Const Message_RecreateWorkbookInstruction = "Recreate Workbook Instruction"
Public Const Message_UpdatedByAnotherUser = "Updated By Another User"
Public Const Message_MissingData = "Missing Data"
Public Const Message_MissingDataAdded = "Missing Data Added"
Public Const Message_RecreateWorkbookAndCopyInstruction = "Recreate Workbook And Copy Data Instruction"

Public Const MessageArg_UpdatedBy = "{Updated By}"



