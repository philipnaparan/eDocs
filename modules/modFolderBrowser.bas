Attribute VB_Name = "modFolderBrowser"

Option Explicit


Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Type Locating
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As Locating) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowserFolder(hwndOwner As Long, promptString As String) As String

     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim spath As String
     Dim udtBI As Locating

     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(promptString, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

     lpIDList = SHBrowseForFolder(udtBI)
     
     If lpIDList Then
        spath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, spath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(spath, vbNullChar)
        If iNull Then spath = Left$(spath, iNull - 1)
     End If

     BrowserFolder = spath
End Function


