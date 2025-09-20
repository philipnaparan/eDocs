Attribute VB_Name = "modPublicVar"
Option Explicit



Public AppDBType As EnumDBType
Public AppConnectionString As String
Public AppCurrentUser As TypeUserInfo

Public LastGenericText As String
Public LastUseFileId As String
Public LastUseFileNamePath As String
Public LastViewedReminderOption As String
Public LastOpenOption As String

'Recently used recordset
Public LastRecordsetA As recordset
Public LastRecordsetB As recordset
Public LastRecordsetC As recordset
 

Public CurrentFolderAccess As TypeFolderRestriction
