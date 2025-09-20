Attribute VB_Name = "modAppFunctions"
Option Explicit

'Function that will format return a generated id
Public Function GeneratedId(ByVal strNumber As String, ByVal strPrefix As String, ByVal strCover As String) As String
    If Len(strCover) <= Len(strNumber) Then
        GeneratedId = strPrefix & strNumber
    Else
        GeneratedId = strPrefix & Left$(strCover, Len(strCover) - Len(strNumber)) & strNumber
    End If
End Function

Public Function getNumberValue(objSource As Object) As Double
    On Error Resume Next
    Dim dblRetVal As Double
    
    dblRetVal = ToNumber(objSource)
    
    getNumberValue = dblRetVal
End Function


Public Function getPercentage(ByVal dNumber As Double, ByVal dPercentage As Double) As Double
    getPercentage = dNumber * (dPercentage / 100)
End Function

Public Function GetDataSource() As String
    On Error Resume Next
    If AppConnectionString <> "" Then
        'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\@Damester\Codes\DB\MasterFile.mdb;Persist Security Info=False
        Dim DSLoc As Integer
        Dim i As Integer
        Dim StartGetDS As Boolean
        Dim Temp As String
        
        DSLoc = InStr(1, AppConnectionString, "Data Source=")
        For i = DSLoc To Len(AppConnectionString)
            If StartGetDS = True Then
                If Mid$(AppConnectionString, i, 1) = ";" Then Exit For
                Temp = Temp & Mid$(AppConnectionString, i, 1)
            Else
                If Mid$(AppConnectionString, i, 1) = "=" Then StartGetDS = True
            End If
        Next i
        
        GetDataSource = Temp
   End If
End Function

Public Function ChangeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": ChangeYNValue = "1"
        Case "N": ChangeYNValue = "0"
        Case "1": ChangeYNValue = "Y"
        Case "0": ChangeYNValue = "N"
    End Select
End Function

'Function used to check if the Ascii is a number or not (return 0 if number)
Public Function IsNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        IsNumber = 0
    Else
        IsNumber = sKeyAscii
    End If
End Function

'Function that will return a currenct format
Public Function ToMoney(ByVal srcCurr As String) As String
   ToMoney = Format$(srcCurr, "#,##0.00")
End Function

'Convert string to number
'I create this istead of ToNumber() co'z val return incorrect value
'ex. Try to see the output of ToNumber("3,800")
'It did not support characters like , and etc.
Public Function ToNumber(ByVal srcCurrency As Variant, Optional RetZeroIfNegative As Boolean, Optional ReturnAsRounded As Boolean) As Double
    If srcCurrency = Null Or srcCurrency = "" Then
        ToNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        
        If ReturnAsRounded = True Then retValue = NumRound(retValue)
        ToNumber = retValue
        retValue = 0
    End If
End Function

'Round a number
Public Function NumRound(Number)
    If Number - Fix(Number) = 0.5 Then Number = Number + 0.01
    NumRound = Round(Number)
End Function




Public Function IsDBValid(ByVal strSQL As String) As Integer
    On Error GoTo err
    
    Dim cn As New Connection
    Dim rs As New recordset
    Dim iRetValue As Integer
    
    cn.Open strSQL
    
    rs.Open "SELECT * FROM tbl_DB_IDENTITY", cn, adOpenStatic, adLockOptimistic
    
    If rs.Fields("IdentityName") = "893JKL3N4690128EJKGRSYS6874*%$^#$%65321A1S4665" Then
        iRetValue = 1
    End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    
    IsDBValid = iRetValue
    
    Exit Function
    
err:
    'Connection error
    If err.Number = -2147467259 Then
        IsDBValid = -1
    'Recordset error
    ElseIf err.Number = -2147467259 Then
        IsDBValid = 0
    End If
End Function

'Procedure used to custom move the recordset cursor
Public Function IsRecExistInRs(ByRef sRS As recordset, ByVal isNum As Boolean, ByVal findStr As String, ByVal sField As String) As Boolean
    Dim bRetVal As Boolean
    
    On Error Resume Next
    
    If sRS.RecordCount < 1 Then Exit Function
    Dim old_pos As Long
    old_pos = sRS.AbsolutePosition
    
    sRS.MoveFirst
    If isNum = True Then
        sRS.Find sField & " = " & findStr
    Else
        sRS.Find sField & " = '" & findStr & "'"
    End If
    If Not sRS.EOF Then bRetVal = True
    sRS.AbsolutePosition = old_pos
    
    IsRecExistInRs = bRetVal
End Function

Public Function RemoveInvalidChar(ByVal SourceText As String, ByVal CharacterToRemove As String) As String
    Dim i As Integer
    Dim strTemp As String
    strTemp = SourceText
    
    For i = 1 To Len(CharacterToRemove)
        strTemp = Replace(strTemp, Mid$(CharacterToRemove, i, 1), "")
    Next i
    
    RemoveInvalidChar = strTemp
    
End Function


Public Function IsNodeKeyExist(ByRef SrcNodes As Nodes, ByVal KeyToFind As String, Optional ExemptedKey As String) As Boolean
    Dim bRetVal As Boolean
    Dim i As Integer
    
    bRetVal = False
    
    For i = 1 To SrcNodes.Count - 1
        If SrcNodes.item(i).key = KeyToFind And SrcNodes.item(i).key <> ExemptedKey Then
            bRetVal = True
            Exit For
        End If
    Next i
    
    IsNodeKeyExist = bRetVal
End Function



Public Function IsNodePathExist(ByRef SrcNodes As Nodes, ByVal PathToFind As String, Optional ExemptedPath As String) As Boolean
    Dim bRetVal As Boolean
    Dim i As Integer
    
    bRetVal = False
    
    For i = 1 To SrcNodes.Count - 1
        If SrcNodes.item(i).FullPath = PathToFind And SrcNodes.item(i).FullPath <> ExemptedPath Then
            bRetVal = True
            Exit For
        End If
    Next i
    
    IsNodePathExist = bRetVal
End Function

'Function used to format recordset
Public Function FormatRecord(ByVal srcField As Field, Optional AllowNewLine As Boolean, Optional BooleanField As String) As String
On Error GoTo err
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        ElseIf srcField.Type = adBoolean Or srcField.Name = BooleanField Then
            If srcField = 0 Then
                strRet = "No"
            Else
                strRet = "Yes"
            End If
        Else
            strRet = srcField
        End If
    End With
    
    FormatRecord = strRet
    
    strRet = vbNullString
err:

End Function


Public Function GetFileNameFromPath(ByVal strFileName As String) As String
    Dim strFName As String
    Dim i As Integer
    
    For i = Len(strFileName) To 1 Step -1
        If Mid$(strFileName, i, 1) = "\" Or Mid$(strFileName, i, 1) = "/" Then Exit For
        strFName = Mid$(strFileName, i, 1) & strFName
    Next i
    
    GetFileNameFromPath = strFName
End Function


Public Function GetNameFromFileName(ByVal strFileName As String) As String
    Dim strFName As String
    Dim i As Integer
    
    For i = 1 To Len(strFileName)
        If Mid$(strFileName, i, 1) = "." Then Exit For
        strFName = strFName & Mid$(strFileName, i, 1)
    Next i
    
    GetNameFromFileName = strFName
End Function

Public Function GetFileExt(ByVal strFileName As String) As String
    Dim strFName As String
    Dim i As Integer
    
    For i = Len(strFileName) To 1 Step -1
        If Mid$(strFileName, i, 1) = "." Then Exit For
        strFName = Mid$(strFileName, i, 1) & strFName
    Next i
    
    GetFileExt = strFName
End Function

Public Function IsExistInCombo(ByVal ItemName As String, ByRef ComboCtrl As ComboBox) As Boolean
    On Error GoTo err
    If ComboCtrl.ListCount > 0 Then
        Dim i As Integer
        For i = 0 To ComboCtrl.ListCount - 1
            If ComboCtrl.List(i) = ItemName Then
                IsExistInCombo = True
                Exit Function
            End If
        Next i
    End If
    
err:
End Function

Public Function FillStr(ByVal StrText As String, ByVal NoOfRepeatation As Long) As String
    If NoOfRepeatation = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To NoOfRepeatation
        FillStr = FillStr & StrText
    Next i
End Function


