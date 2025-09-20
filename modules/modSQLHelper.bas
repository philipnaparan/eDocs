Attribute VB_Name = "modSQLHelper"
Option Explicit

Public LAST_GENERATED_IDENTITY As Long

'****************************************
'GetRecords
'****************************************
'[Description]
'Use to get records
'--------------------
'[Parameter]
'strSQL - SQL Statement use to retrieve records
'cnConnection - Source connection
'--------------------
'[Return]
'Return a Recordset
'****************************************
Public Function GetRecords(ByVal strSQL As String, Optional cnConnection As Connection) As recordset
    On Error GoTo err
    
    'Recordset to return
    Dim rsReturn As recordset
    Dim cn As Connection
    
    If cnConnection Is Nothing Then
        Set cn = New Connection
        cn.CursorLocation = adUseClient
        cn.Open AppConnectionString
    Else
        Set cn = cnConnection
    End If
    
    Set rsReturn = New recordset
    
    rsReturn.CursorLocation = adUseClient
    rsReturn.CursorType = adOpenStatic
    rsReturn.LockType = adLockOptimistic
    
    rsReturn.Open strSQL, cn
       
    Set rsReturn.ActiveConnection = Nothing
    If cnConnection Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
        
    Set GetRecords = rsReturn
    
    Exit Function
err:
    MsgBox err.Description
    Set GetRecords = Nothing
End Function

'Public Function GetCommand(ByVal strSQL As String, Optional cnConnection As Connection) As Command
'    On Error GoTo err
'
'    'Recordset to return
'    Dim cmdReturn As Command
'    Dim cn As Connection
'
'    If cnConnection Is Nothing Then
'        Set cn = New Connection
'        cn.CursorLocation = adUseClient
'        cn.Open AppConnectionString
'    Else
'        Set cn = cnConnection
'    End If
'
'    Set cmdReturn = New Command
'
'    cmdReturn.
'
'    cmdReturn.CursorLocation = adUseClient
'    cmdReturn.CursorType = adOpenStatic
'    cmdReturn.LockType = adLockOptimistic
'
'    cmdReturn.Open strSQL, cn
'
'    Set cmdReturn.ActiveConnection = Nothing
'    If cnConnection Is Nothing Then
'        If cn.State = adStateOpen Then cn.Close
'        Set cn = Nothing
'    End If
'
'    Set GetCommand = cmdReturn
'
'    Exit Function
'err:
'    MsgBox err.Description
'    Set GetCommand = Nothing
'End Function

Public Function ExecCommand(ByVal CommandText As String, _
                            ByRef CommandParam() As TypeCommandParam, _
                            ByRef RecordStructure As recordset, _
                            Optional cnConnection As Connection) As Boolean
    'On Error GoTo err
    
    Dim bReturn As Boolean
    Dim cmd As Command
    Dim param As Parameter
    Dim cn As Connection
    
    If cnConnection Is Nothing Then
        Set cn = New Connection
        cn.CursorLocation = adUseClient
        cn.Open AppConnectionString
    Else
        Set cn = cnConnection
    End If
    
    bReturn = False
    
    Set cmd = New Command
    cmd.CommandText = CommandText
    Set cmd.ActiveConnection = cn
    
    Dim i As Long
    For i = LBound(CommandParam) To UBound(CommandParam)
        Set param = cmd.CreateParameter
        param.Direction = adParamInput
        param.Name = CommandParam(i).ParamName
        param.Value = CommandParam(i).ParamValue
        param.Type = RecordStructure.Fields(Replace(CommandParam(i).ParamName, "@", "")).Type
        MsgBox RecordStructure.Source

        cmd.Parameters.Append param
    Next i
    
    cmd.Execute
    bReturn = True
    
    If cnConnection Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Set cmd = Nothing
    Set param = Nothing
        
    ExecCommand = bReturn
    
    Exit Function
err:
    MsgBox err.Description
End Function

Public Function GetRecordStructure(ByVal TableName As String) As recordset
    Set GetRecordStructure = GetRecords("SELECT * FROM [" & TableName & "] WHERE 1=0")
End Function


Public Function GetRowPK(ByVal strTableName As String) As Long
    On Error Resume Next
    
    Dim cn As New Connection
    Dim rs As New recordset
    
    Dim lCurrentIndex As Long
    
    cn.Open AppConnectionString
    rs.Open "SELECT * FROM t_DB_PRIMARY_KEY_GENERATOR WHERE TableName='" & strTableName & "'", cn, adOpenStatic, adLockOptimistic
    
    lCurrentIndex = rs.Fields("NextIndex")
    
    rs.Fields("NextIndex") = lCurrentIndex + 1
    rs.Update
    
    If rs.State = adStateOpen Then rs.Close
    If cn.State = adStateOpen Then cn.Close
        
    Set rs = Nothing
    Set cn = Nothing
    
    GetRowPK = lCurrentIndex
End Function

Public Function GetRowId(ByVal strTableName As String) As String
    On Error Resume Next
    
    Dim cn As New Connection
    Dim rs As New recordset
    
    Dim lCurrentIndex As Long
    Dim strGeneratedId As String
    
    Dim strPrefix As String
    Dim strCover As String
    
    cn.Open AppConnectionString
    rs.Open "SELECT * FROM t_DB_ID_GENERATOR WHERE TableName='" & strTableName & "'", cn, adOpenStatic, adLockOptimistic
    
    strPrefix = rs.Fields("Prefix")
    strCover = rs.Fields("Cover")
    
    lCurrentIndex = rs.Fields("NextIndex")
    strGeneratedId = GeneratedId(rs.Fields("NextIndex"), strPrefix, strCover)
    
    rs.Fields("NextIndex") = lCurrentIndex + 1
    rs.Update
    
    If rs.State = adStateOpen Then rs.Close
    If cn.State = adStateOpen Then cn.Close
        
    Set rs = Nothing
    Set cn = Nothing
    
    GetRowId = strGeneratedId
End Function

Public Function GetConcurrencyId(ByVal strTableName As String) As Long
    On Error Resume Next
    
    Dim cn As New Connection
    Dim rs As New recordset
    
    Dim lCurrentIndex As Long
    
    cn.Open AppConnectionString
    rs.Open "SELECT * FROM t_DB_CURRENCY_GENERATOR WHERE TableName='" & strTableName & "'", cn, adOpenStatic, adLockOptimistic
    
    lCurrentIndex = rs.Fields("NextIndex")
    
    rs.Fields("NextIndex") = lCurrentIndex + 1
    rs.Update
    
    If rs.State = adStateOpen Then rs.Close
    If cn.State = adStateOpen Then cn.Close
        
    Set rs = Nothing
    Set cn = Nothing
    
    GetConcurrencyId = lCurrentIndex
End Function

'****************************************
'GetRecordCount
'****************************************
'[Description]
'Use to get the record count.
'--------------------
'[Parameter]
'StrSQL - SQL Statement use to retrieve records
'conn - Reference connection
'--------------------
'[Return]
'Return a long total numbers of records
'****************************************
Public Function GetRecordCount(ByVal strSQL As String, Optional cnConnection As Connection) As Long
    On Error Resume Next
    
    Dim lCount As Long
    Dim rs As recordset
    Dim cn As Connection
        
    lCount = 0
    If cnConnection Is Nothing Then
        Set cn = New Connection
        cn.CursorLocation = adUseClient
        cn.Open AppConnectionString
    Else
        Set cn = cnConnection
    End If

    
    Set rs = New recordset
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    lCount = rs.RecordCount
    
     rs.Close
    Set rs = Nothing
    
     If cnConnection Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
   
    GetRecordCount = lCount
    
End Function

Public Function DeleteRecord(ByVal strSQL As String, Optional cnConnection As Connection) As Long
    On Error GoTo err
        Dim lResult As Long
        Dim cn As Connection
        
        lResult = 0
        If cnConnection Is Nothing Then
            Set cn = New Connection
            cn.Open AppConnectionString
        Else
            Set cn = cnConnection
        End If
        
        cn.Execute strSQL
        lResult = 1
        
        If cnConnection Is Nothing Then
            If cn.State = adStateOpen Then cn.Close
            Set cn = Nothing
        End If
        
        DeleteRecord = lResult
    Exit Function
err:
    DeleteRecord = err.Number
End Function

Public Function SaveRecord(ByVal strSQLSaveIn As String, ByRef rsRecord As recordset, Optional cnConnection As Connection, Optional IsUpdationMode As Boolean, Optional strIdentityTable As String, Optional ByPassNullValue As Boolean) As Long
    On Error GoTo err
        Dim lResult As Long
        Dim cn As Connection
        Dim rsSave As recordset
        
        lResult = 0
        If cnConnection Is Nothing Then
            Set cn = New Connection
            cn.Open AppConnectionString
        Else
            Set cn = cnConnection
        End If
        
        If strSQLSaveIn = "" Then strSQLSaveIn = rsRecord.Source
        
        Dim iCurrField As Integer
        Set rsSave = New recordset
        rsSave.Open strSQLSaveIn, cn, adOpenStatic, adLockOptimistic
        
        
        'Save the record
        If IsUpdationMode = False Then
            rsSave.AddNew
        Else
            If rsSave.RecordCount < 1 Then SaveRecord = 0: Exit Function
        End If
        
        For iCurrField = 0 To rsSave.Fields.Count - 1
            If ByPassNullValue = True Then
                If IsNull(rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value) = False Then
                    'MsgBox rsSave.Fields(iCurrField).Name & "=" & rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value
                    rsSave.Fields(rsSave.Fields(iCurrField).Name) = rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value
                End If
            Else
                rsSave.Fields(rsSave.Fields(iCurrField).Name) = rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value
            End If
        Next iCurrField
        rsSave.Update
        
        
        If IsUpdationMode = False And strIdentityTable <> "" Then
            Dim rsIdentity As recordset
            
            Set rsIdentity = New recordset
            
            If AppDBType = adDBTypeMSAccess Then
                rsIdentity.Open "SELECT @@IDENTITY as LastIdentity FROM " & strIdentityTable, cn
            Else
                rsIdentity.Open "SELECT IDENT_CURRENT('" & strIdentityTable & "') as LastIdentity", cn
            End If
            LAST_GENERATED_IDENTITY = Val(rsIdentity.Fields("LastIdentity"))
            
            If Not rsIdentity Is Nothing Then
                If rsIdentity.State = adStateOpen Then rsIdentity.Close
                Set rsIdentity = Nothing
            End If
        End If
        
        If Not cnConnection Is Nothing Then
            If cn.State = adStateOpen Then cn.Close
            Set cn = Nothing
        End If
        
        lResult = 1
        SaveRecord = lResult
    Exit Function
err:
    If err.Number = -2147217887 Then Resume Next
    SaveRecord = err.Number
End Function


'Public Function SaveRecord(ByVal strSQLSaveIn As String, ByRef rsRecord As Recordset, Optional cnConnection As Connection, Optional IsUpdationMode As Boolean) As Long
'    'On Error GoTo err
'        Dim lResult As Long
'        Dim cn As Connection
'        Dim rsSave As Recordset
'
'        lResult = 0
'        If cnConnection Is Nothing Then
'            Set cn = New Connection
'            cn.Open AppConnectionString
'        Else
'            Set cn = cnConnection
'        End If
'
'        Dim iCurrField As Integer
'
'        Set rsSave = New Recordset
'        rsSave.Open strSQLSaveIn, cn, adOpenStatic, adLockOptimistic
'
'        'Save the record
'        If IsUpdationMode = False Then rsSave.AddNew
'            For iCurrField = 0 To rsSave.Fields.Count - 1
'                'MsgBox rsSave.Fields(iCurrField).Name & " = " & rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value
'                rsSave.Fields(rsSave.Fields(iCurrField).Name) = rsRecord.Fields(rsSave.Fields(iCurrField).Name).Value
'            Next iCurrField
'        rsSave.Update
'        lResult = 1
'
'        If cnConnection Is Nothing Then
'            If cn.State = adStateOpen Then cn.Close
'            Set cn = Nothing
'        End If
'
'        SaveRecord = lResult
'    Exit Function
'err:
'    SaveRecord = err.Number
'End Function

Public Function ConcurrencyInfo(ByVal strSQL As String, ByVal lConcurrencyValue As Long) As Integer
    On Error Resume Next
    
    Dim cn As New Connection
    Dim rs As New recordset
    
    Dim iResult As Integer
    
    cn.CursorLocation = adUseClient
    cn.Open AppConnectionString
    rs.Open strSQL, cn, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        If ToNumber(rs.Fields("ConcurrencyId")) = lConcurrencyValue Then
            iResult = 1
        Else
            iResult = 0
        End If
    Else
        'Means the record is no longer exist1
        iResult = -1
    End If
    
    If rs.State = adStateOpen Then rs.Close
    If cn.State = adStateOpen Then cn.Close
        
    Set rs = Nothing
    Set cn = Nothing
    
    ConcurrencyInfo = iResult
End Function

'Function that return the value of a certain field
Public Function GetValueAt(ByVal strSQL As String, ByVal strWhichField As String, Optional AppConnectionStringString As String) As String
    On Error Resume Next
    
    Dim rs As New recordset
    Dim cn As New Connection
    
    If AppConnectionStringString = "" Then AppConnectionStringString = AppConnectionString
    cn.Open AppConnectionStringString
    
    rs.CursorLocation = adUseClient
    rs.Open strSQL, cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then GetValueAt = rs.Fields(strWhichField)
    
    Set rs = Nothing
    Set cn = Nothing
End Function

Public Sub CopyRecordsetFields(ByRef rsSource As recordset, ByRef rsDestination As recordset)
     Dim iCurrField As Integer

    'Copy the fields
    For iCurrField = 0 To rsSource.Fields.Count - 1
        rsDestination.Fields.Append rsSource.Fields(iCurrField).Name, rsSource.Fields(iCurrField).Type, rsSource.Fields(iCurrField).DefinedSize, rsSource.Fields(iCurrField).Attributes
    Next iCurrField
End Sub

Public Sub CopyCurrentRecordsetRecordValue(ByRef rsSource As recordset, ByRef rsDestination As recordset)
     Dim iCurrField As Integer

    rsDestination.AddNew
        'Copy the current value
        For iCurrField = 0 To rsDestination.Fields.Count - 1
            rsDestination.Fields(rsDestination.Fields(iCurrField).Name) = rsSource.Fields(rsDestination.Fields(iCurrField).Name).Value
        Next iCurrField
    rsDestination.Update
End Sub

Public Sub SyncronizeRecordsetBinding(ByRef rsSource As recordset, ByRef frmSourceFormContainer As Form)
    On Error GoTo err
    Dim ctrlFormControl As Control
    For Each ctrlFormControl In frmSourceFormContainer.Controls
        If ctrlFormControl.DataField <> "" Then
            'Control that having a text value
            If (TypeOf ctrlFormControl Is TextBox Or TypeOf ctrlFormControl Is ComboBox) Then rsSource.Fields(ctrlFormControl.DataField) = ctrlFormControl.Text
            'Control that having a value property
            If (TypeOf ctrlFormControl Is DTPicker) Then rsSource.Fields(ctrlFormControl.DataField) = ctrlFormControl.Value
        End If
    Next ctrlFormControl
    Set ctrlFormControl = Nothing
    Exit Sub
err:
    If err.Number = 438 Then Resume Next 'DataField is not supported error
    If err.Number = -2147217887 Then Resume Next 'Multiple step error
    If err.Number = 3265 Then Resume Next 'Cannot find name ordinal
    InputBox err.Description, "ERROR", err.Number
End Sub

Public Sub AddField(ByVal strAddToTableName As String, ByVal strNewFieldName As String, Optional AppConnectionStringString As String)
    On Error GoTo err
    
    Dim rs As New recordset
    Dim cn As New Connection
    
    If AppConnectionStringString = "" Then AppConnectionStringString = AppConnectionString
    cn.Open AppConnectionStringString
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM " & strAddToTableName, cn, adOpenStatic, adLockOptimistic
    
    'Add a new field if false
    If IsFieldNameExist(rs, strNewFieldName) = False Then
        rs.Fields.Append strNewFieldName, adVarNumeric
    End If
    
    Set rs = Nothing
    Set cn = Nothing
    
err:
    Exit Sub
End Sub


Private Function IsFieldNameExist(ByRef rsFieldContainer As recordset, ByVal strFieldNameToCheck As String) As Boolean
    On Error GoTo err
    If rsFieldContainer.Fields(strFieldNameToCheck).Name <> "" Then IsFieldNameExist = True
    
err:
    IsFieldNameExist = False
End Function


Public Function GetSettings(ByVal SettingsName As String) As String
    On Error GoTo err
    
    'Recordset to return
    Dim strReturnString As String
    Dim cn As Connection
    Dim rs As recordset
    
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\settings.mdb;Persist Security Info=False;Jet OLEDB:Database Password=pay123"
    
    Set rs = GetRecords("SELECT * FROM [Settings] WHERE [SettingsName]='" & SettingsName & "'", cn)
    
    If Not rs Is Nothing Then
        strReturnString = rs.Fields("SettingsValue")
        
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    
       
    GetSettings = strReturnString
    
    Exit Function
err:
    GetSettings = vbNullString
End Function


Public Function SaveSettings(ByVal SettingsName As String, ByVal NewValue As String) As Boolean
    On Error GoTo err
    
       
    SaveSettings = ExecSQL("UPDATE [Settings] SET [SettingsValue]='" & NewValue & "' WHERE [SettingsName]='" & SettingsName & "'", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\settings.mdb;Persist Security Info=False;Jet OLEDB:Database Password=pay123")
    
    Exit Function
err:
    SaveSettings = False
End Function

Public Function ExecSQL(ByVal sqlCommand As String, Optional ConnectionString As String) As Boolean
    On Error GoTo err
    
    sqlCommand = Replace(sqlCommand, "'''", "''")
    
    'Recordset to return
    Dim bIsSuccess As Boolean
    Dim cn As Connection
    Dim strConn As String
    
    Set cn = New Connection
    cn.CursorLocation = adUseServer
    
    If ConnectionString = "" Then
        strConn = AppConnectionString
    Else
        strConn = ConnectionString
    End If
    cn.Open strConn
    
    bIsSuccess = False
    cn.Execute sqlCommand
    bIsSuccess = True
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
       
    ExecSQL = bIsSuccess
    
    Exit Function
err:
    MsgBox err.Description
    ExecSQL = False
End Function

Public Function CreateConnection(Optional ConnectionString As String) As Connection
    On Error GoTo err
    
    'Connection to return
    Dim cn As Connection
    
    If ConnectionString = "" Then ConnectionString = AppConnectionString
        
    Set cn = New Connection
    cn.CursorLocation = adUseClient
    cn.Open AppConnectionString
    
    Set CreateConnection = cn
    
    Exit Function
err:
    MsgBox err.Description
    Set CreateConnection = Nothing
End Function




'**************************************************************
'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME
'
'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACCESS)
'
'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE
'
'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MYTABLE", oConn, adOpenKeyset, _
   adLockOptimistic
'oRs.AddNew
'
'SaveFileToDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
'oRs.Update
'oRs.Close
'**************************************************************

Public Function SaveFileToDB(ByVal fileName As String, rs As recordset, FieldName As String) As Boolean

        Dim iFileNum As Integer
        Dim lFileLength As Long
        
        Dim abBytes() As Byte
        Dim iCtr As Integer
        
        On Error GoTo ErrorHandler
        If Dir(fileName) = "" Then Exit Function
        If Not TypeOf rs Is ADODB.recordset Then Exit Function
        
        'read file contents to byte array
        iFileNum = FreeFile
        Open fileName For Binary Access Read As #iFileNum
        lFileLength = LOF(iFileNum)
        ReDim abBytes(lFileLength)
        Get #iFileNum, , abBytes()
        
        'put byte array contents into db field
        rs.Fields(FieldName).AppendChunk abBytes()
        Close #iFileNum
        
        SaveFileToDB = True
ErrorHandler:
End Function


'************************************************
'PURPOSE: LOADS BINARY DATA IN RECORDSET RS,
'FIELD FieldName TO a File Named by the FileName parameter

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MyTable", oConn, adOpenKeyset,
' adLockOptimistic
'DownloadFileFromDB "C:\MyDocuments\MyDoc.Doc",  oRs, "MyFieldName"
'oRs.Close
'************************************************
Public Function DownloadFileFromDB(fileName As String, rs As recordset, FieldName As String) As Boolean
        Dim iFileNum As Integer
        Dim lFileLength As Long
        Dim abBytes() As Byte
        Dim iCtr As Integer
        
        On Error GoTo ErrorHandler
        If Not TypeOf rs Is ADODB.recordset Then Exit Function
        
        iFileNum = FreeFile
        Open fileName For Binary As #iFileNum
        lFileLength = LenB(rs(FieldName))
        
        abBytes = rs(FieldName).GetChunk(lFileLength)
        Put #iFileNum, , abBytes()
        Close #iFileNum
        DownloadFileFromDB = True
        
ErrorHandler:
End Function


Public Function LoadFileImageDB(rs As recordset, FieldName As String) As IPictureDisp
        Dim strFileName As String
        Dim picRetImage As IPictureDisp
        Dim iFileNum As Integer
        Dim lFileLength As Long
        Dim abBytes() As Byte
        Dim iCtr As Integer
        
        'On Error GoTo ErrorHandler
        If Not TypeOf rs Is ADODB.recordset Then Exit Function
        If rs(FieldName) Is Nothing Then Exit Function
        If GetFileLength(rs(FieldName)) = 0 Then Exit Function
        
        strFileName = App.Path & "\" & GenerateFileName
        iFileNum = FreeFile
        lFileLength = GetFileLength(rs(FieldName))
        
        Open strFileName For Binary As #iFileNum
            abBytes = rs(FieldName).GetChunk(lFileLength)
            Put #iFileNum, , abBytes()
        Close #iFileNum
        
        Set picRetImage = LoadPicture(strFileName)
        KillFile (strFileName)
        Set LoadFileImageDB = picRetImage
        
ErrorHandler:
End Function

Public Function GetFileLength(Expression As Object) As Long
    On Error Resume Next
    GetFileLength = LenB(Expression)
End Function

Public Sub KillFile(strFileName As String)
    On Error Resume Next
    Kill strFileName
End Sub

Public Function GenerateFileName(Optional ExtName As String) As String
    If ExtName = "" Then ExtName = ".tmp"
    Randomize
    GenerateFileName = GenerateFileId & ExtName
End Function

Public Function GenerateFileId() As String
    Randomize
    GenerateFileId = (CLng(Rnd() * 1000000000) & CLng(Rnd() * 1000000000))
End Function


