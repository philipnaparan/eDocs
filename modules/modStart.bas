Attribute VB_Name = "modStart"
Option Explicit


Public LoginOk As Boolean
Public ProductCode As String
Public IsDemo As Boolean

Public Sub Main()
    On Error Resume Next
    
'    SaveSetting "Microsoft VBA Scripter", "Application", "PathInfo", ""
'    SaveSetting "Microsoft VBA Scripter", "Application", "PathCode", ""
'    Exit Sub
    
    AppDBType = adDBTypeSQLServer
    
    If App.PrevInstance = True Then
        MsgBox "Application is currently running.", vbInformation
        Exit Sub
    End If
    
    Dim objExpiry As CExpiry
    Set objExpiry = New CExpiry
'    objExpiry.Reset: Exit Sub
    
    If IsAuthenticated() = False Then
        If objExpiry.IsExpired = True Then
REGISTER:
            'Activate the product
            If IsAuthenticated() = False Then
                frmProdActivation.Show vbModal
                If ProductCode = "" Then
                    Exit Sub
                Else
                    If IsAuthenticated(ProductCode) = False Then
                        MsgBox "Invalid activation code.", vbCritical
                        Exit Sub
                    End If
                End If
            End If
        Else
            frmTrialInfo.Show vbModal
            If LastGenericText = "reg" Then
                GoTo REGISTER
            Else
                IsDemo = True
            End If
        End If
    End If
    
    Set objExpiry = Nothing
    
'    Exit Sub
    frmFileManager.Show
    
    frmSplash.Show vbModal, frmFileManager
    frmLogin.Show vbModal, frmFileManager
    
    If LoginOk = False Then
        Unload frmFileManager
    Else
        With frmFileManager
            .InitializeForm
        End With
    End If
End Sub

Public Function IsAuthenticated(Optional ProdCode As String) As Boolean
    
    On Error GoTo err
    Dim retVal As Boolean
    Dim key As String
    Dim decodedKey As String
    Dim diskSerial As String
    
    If ProdCode = "" Then
        key = GetSetting("Microsoft VBA Scripter", "Application", "PathCode")
        LastGenericText = GetSetting("Microsoft VBA Scripter", "Application", "PathInfo")
    Else
        key = ProdCode
    End If
    
    If key = "" Then
        retVal = False
    Else

        diskSerial = AESDencrypt(GetEncryptedDriveSerial(), EncPass(), enum256Bit)
        decodedKey = AESDencrypt(key, EncPass(), enum256Bit)

        If decodedKey = diskSerial Then
            SaveSetting "Microsoft VBA Scripter", "Application", "PathInfo", LastGenericText
            SaveSetting "Microsoft VBA Scripter", "Application", "PathCode", key
            'SaveSetting "Microsoft VBA Scripter", "Application", "PathData", AESEncrypt(LastGenericText, EncPass, enum256Bit)

            retVal = True
        Else
            retVal = False
        End If
    End If

    'IsAuthenticated = retVal 'disable authentication
    IsAuthenticated = True
    Exit Function
err:
    'IsAuthenticated = False  'disable authentication
    IsAuthenticated = True
End Function

Public Function EncPass() As String
    EncPass = DeCode("6a@pm1vbn13sw$wpl5qwt|pq;&caq@") 'sempron501
End Function

Private Function GetEncryptedDriveSerial() As String
    On Error Resume Next
    GetEncryptedDriveSerial = AESEncrypt(GetDriveSerial(), "sempron501", enum256Bit)
End Function


Public Function GetDriveSerial() As String
    Dim h As CDriveSerial
    
    Dim hT As Long
    Dim uW() As Byte
    Dim dW() As Byte
    Dim pW() As Byte
    Dim retVal As String
   
    Set h = New CDriveSerial
   
    With h
        .CurrentDrive = 0
       
        retVal = .GetSerialNumber
    End With
   
    Set h = Nothing
    
    GetDriveSerial = retVal
End Function

