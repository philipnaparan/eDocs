Attribute VB_Name = "modIOHelper"
Option Explicit

'****************************************
'GetImage
'****************************************
'[Description]
'Function use to load images from the the application directory
'for the application used.
'--------------------
'[Parameter]
'ImageName - Name of the image file to load
'--------------------
'[Return]
'Return an IPictureDisp
'****************************************
Public Function GetImage(ByVal ImageName As String) As IPictureDisp
    On Error Resume Next
    Set GetImage = LoadPicture(App.Path & "\Images\" & ImageName)
End Function



Public Sub LunchFileWithDialog(ByVal fileName As String)
    ShellExecute 0, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " _
        & fileName, "", vbNormalFocus
End Sub

Public Function FileExists(ByVal fileName As String) As Boolean
    On Error GoTo ErrorHandler
    FileExists = (GetAttr(fileName) And vbDirectory) = 0
ErrorHandler:
End Function

Public Function GetFileSizeString(ByVal fileName As String) As String
    On Error GoTo ErrorHandler
    
    GetFileSizeString = GetFileSizeInfo(GetFileSize(fileName))
    
ErrorHandler:
End Function

Public Function GetFileSize(ByVal fileName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim fso As New FileSystemObject
    Dim retVal As Long
    
    retVal = fso.GetFile(fileName).Size
    
    GetFileSize = retVal
    
ErrorHandler:
End Function

Public Function GetFileSizeInfo(ByVal FileSize As Long) As String
    Select Case FileSize
        Case 0 To 999
            GetFileSizeInfo = Round(FileSize, 2) & " Bytes"
            Exit Function
        Case 1000 To 999999
            GetFileSizeInfo = Round(FileSize / 1000, 2) & " KB"
            Exit Function
        Case 1000000 To 999999999
            GetFileSizeInfo = Round(FileSize / 1000000, 2) & " MB"
            Exit Function
        Case Is >= 1000000000
            GetFileSizeInfo = Round(FileSize / 1000000000, 2) & " GB"
            Exit Function
    End Select
End Function


Public Function GetFileNameWithOutExt(ByVal fileName As String)

    GetFileNameWithOutExt = Replace(fileName, "." & GetFileExt(fileName), "")
End Function

