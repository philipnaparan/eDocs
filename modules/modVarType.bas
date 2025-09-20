Attribute VB_Name = "modVarType"

Option Explicit


Public Type TypeUserInfo
    UserName As String
    CompleteName As String
    UserId As Long
    UserGroupId As Long
    bCanViewConfidential As Boolean
    bCanAdd As Boolean
    bCanEdit As Boolean
    bCanChkOut As Boolean
    bCanDelete As Boolean
    bCanImport As Boolean
    bCanExport As Boolean
    bCanAddFolder As Boolean
    bCanEditFolder As Boolean
    bCanDeleteFolder As Boolean
    bCanManageTemplates As Boolean
    bIsSysAdmin As Boolean
End Type

Public Type TypeFolderRestriction
    bDenyFolderAccess As Boolean
    bDenyFolderEdit As Boolean
    bDenyFolderDelete As Boolean
    bDenyOpenFile As Boolean
    bDenyCreateFile As Boolean
    bDenyEditFile As Boolean
    bDenyDeleteFile As Boolean
    bDenyCheckOut As Boolean
    bDenyFileImport As Boolean
    bDenyFileExport As Boolean
    FolderId As Long
End Type

Public Enum EnumDBType
    adDBTypeSQLServer = 0
    adDBTypeMSAccess = 1
End Enum

Public Type TypeCommandParam
    ParamName As String
    ParamValue As Variant
End Type

'Enumerator for form state
Public Enum FormState
        adStateAddMode = 0
        adStateEditMode = 1
        adStateViewMode = 2
        adStatePopupMode = 3
End Enum

