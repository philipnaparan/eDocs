VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Reset()
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataF", ""
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataL", ""
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataE", ""
    SaveSetting "Microsoft VBA Scripter", "Application", "Errors", ""
End Sub

Public Sub SetExpiry()
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataF", AESEncrypt(Date, EncPass, enum256Bit)
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataL", AESEncrypt(Date, EncPass, enum256Bit)
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataE", AESEncrypt(DateAdd("d", 30, Date), EncPass, enum256Bit)
End Sub


Public Function GetExpiry() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting("Microsoft VBA Scripter", "Application", "MetaDataE")
    If retVal <> "" Then retVal = AESDencrypt(retVal, EncPass, enum256Bit)
    
    GetExpiry = retVal
End Function


Public Function GetLastUse() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting("Microsoft VBA Scripter", "Application", "MetaDataL")
    If retVal <> "" Then retVal = AESDencrypt(retVal, EncPass, enum256Bit)
    
    GetLastUse = retVal
End Function

Public Function GetFirstUse() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting("Microsoft VBA Scripter", "Application", "MetaDataF")
    If retVal <> "" Then retVal = AESDencrypt(retVal, EncPass, enum256Bit)
    
    GetFirstUse = retVal
End Function

Public Function GetExpired() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting("Microsoft VBA Scripter", "Application", "Errors")
    If retVal <> "" Then retVal = "Yes"

    GetExpired = retVal
End Function

Public Sub SetExpired()
    SaveSetting "Microsoft VBA Scripter", "Application", "Errors", AESEncrypt("Yes", EncPass, enum256Bit)
End Sub


Public Function IsExpired() As Boolean
    On Error Resume Next
    Dim retVal As Boolean
    
    If GetExpiry = "" And GetLastUse = "" And GetFirstUse = "" Then
        SetExpiry
        IsExpired = False
        Exit Function
    End If
    
    If GetExpired = "Yes" Then
        IsExpired = True
        Exit Function
    End If
    
    If CLng(Date) > CLng(CDate(GetExpiry)) Then
        SetExpired
        UpdateLastUse
        retVal = True
'    ElseIf CLng(CDate(GetLastUse)) > CLng(CDate(GetExpiry)) Then
'        SetExpired
'        UpdateLastUse
'        retVal = True
    ElseIf CLng(Date) < CLng(CDate(GetLastUse)) Then
        Date = CDate(GetLastUse)
    Else
        UpdateLastUse
    End If
    
    IsExpired = retVal
End Function


Public Sub UpdateLastUse()
    On Error Resume Next
    SaveSetting "Microsoft VBA Scripter", "Application", "MetaDataL", AESEncrypt(Date, EncPass, enum256Bit)
End Sub
