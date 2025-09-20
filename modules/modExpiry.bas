Attribute VB_Name = "modExpiry"
Option Explicit


Public Sub SetExpiry()
    SaveSetting App.EXEName, "Application", "MetaDataF", AESEncrypt(Now, EncPass, enum256Bit)
    SaveSetting App.EXEName, "Application", "MetaDataL", AESEncrypt(Now, EncPass, enum256Bit)
    SaveSetting App.EXEName, "Application", "MetaDataE", AESEncrypt(DateAdd("d", 30, Now), EncPass, enum256Bit)
End Sub


Public Function GetExpiry() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting(App.EXEName, "Application", "MetaDataE")
    
    GetExpiry = retVal
End Function


Public Function GetLastDemoUse() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting(App.EXEName, "Application", "MetaDataL")
    
    GetExpiry = retVal
End Function

Public Function GetExpiry() As String
    On Error Resume Next
    Dim retVal As String
    
    retVal = GetSetting(App.EXEName, "Application", "MetaDataE")
    
    GetExpiry = retVal
End Function



