Attribute VB_Name = "modAppProcedure"
Option Explicit

Public Sub BindControlToRS(ByRef objFormContainer As Form, ByRef objRs As recordset)
        
    On Error Resume Next
    Dim ctrlControl As Control
    For Each ctrlControl In objFormContainer.Controls
        If ctrlControl.Tag <> "NoBind" Then
            If ctrlControl.DataField <> "" Then
                Set ctrlControl.DataSource = objRs
                'Set ctrlControl.RowSource = objRs
            End If
        End If
    Next ctrlControl
    Set ctrlControl = Nothing
End Sub

Public Sub UnBindControl(ByRef objFormContainer As Form)
        
    On Error Resume Next
    Dim ctrlControl As Control
    For Each ctrlControl In objFormContainer.Controls
        If ctrlControl.Tag <> "NoBind" Then
            If ctrlControl.DataField <> "" Then
                Set ctrlControl.DataSource = Nothing
                'Set ctrlControl.RowSource = Nothing
            End If
        End If
    Next ctrlControl
    Set ctrlControl = Nothing
End Sub


Public Sub EnableBindedControls(ByRef objFormContainer As Form, ByVal isEnable As Boolean)
        
    On Error Resume Next
    Dim ctrlControl As Control
    For Each ctrlControl In objFormContainer.Controls
        If ctrlControl.Tag <> "NoBind" Then
            If ctrlControl.DataField <> "" Then
                Set ctrlControl.Enabled = isEnable
            End If
        End If
    Next ctrlControl
    Set ctrlControl = Nothing
End Sub




'Procedure used to custom move the recordset cursor
Public Sub MoveRs(ByRef sRS As recordset, ByVal isNum As Boolean, ByVal findStr As String, ByVal sField As String)
    If sRS.RecordCount < 1 Then Exit Sub
    Dim old_pos As Long
    sRS.MoveFirst
    old_pos = sRS.AbsolutePosition
    If isNum = True Then
        sRS.Find sField & " = " & findStr
    Else
        sRS.Find sField & " = '" & findStr & "'"
    End If
    If sRS.EOF Then sRS.AbsolutePosition = old_pos
    old_pos = 0
End Sub


'Procedure used to clear the text content
Public Sub ClearTextBox(ByRef sForm As Form)
    Dim ctrlFormControl As Control
    For Each ctrlFormControl In sForm.Controls
        If (TypeOf ctrlFormControl Is TextBox) Then
            If ctrlFormControl.TabStop = True Then ctrlFormControl = vbNullString
        End If
    Next ctrlFormControl
    Set ctrlFormControl = Nothing
End Sub

'Procedure used to clear the text content
Public Sub ClearText(ByRef sForm As Form)
    On Error Resume Next
    Dim ctrlFormControl As Control
    For Each ctrlFormControl In sForm.Controls
        If Not (TypeOf ctrlFormControl Is Label) Then
            If ctrlFormControl.TabStop = True Then ctrlFormControl = vbNullString
        End If
    Next ctrlFormControl
    Set ctrlFormControl = Nothing
End Sub

'Procedure used to remove invalid string
Public Sub RemoveInvalidString(ByRef SourceForm As Form, ByVal TextToRemove As String, _
                                Optional TextToRemove1 As String, _
                                Optional TextToRemove2 As String, _
                                Optional TextToRemove3 As String, _
                                Optional TextToRemove4 As String, _
                                Optional TextToRemove5 As String)
    On Error Resume Next
    Dim ctrlFormControl As Control
    For Each ctrlFormControl In SourceForm.Controls
        If Not (TypeOf ctrlFormControl Is Label) Then
            If ctrlFormControl.TabStop = True Then
                If TextToRemove <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove, "")
                If TextToRemove1 <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove1, "")
                If TextToRemove2 <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove2, "")
                If TextToRemove3 <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove3, "")
                If TextToRemove4 <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove4, "")
                If TextToRemove5 <> "" Then ctrlFormControl.Text = Replace(ctrlFormControl.Text, TextToRemove5, "")
            End If
        End If
    Next ctrlFormControl
    Set ctrlFormControl = Nothing
End Sub

'Procedure used to locked input field
Public Sub LockControl(ByRef sForm As Form, ByVal bLockControl As Boolean)
    On Error Resume Next
    Dim ctrlFormControl As Control
    For Each ctrlFormControl In sForm.Controls
        If ctrlFormControl.Tag <> "NoLock" Then ctrlFormControl.Locked = bLockControl
    Next ctrlFormControl
    Set ctrlFormControl = Nothing
End Sub



'Procedure used to promp unexpected errors
Public Sub PromptError(ByVal objError As ErrObject, ByVal strModuleName As String, ByVal strOccurIn As String)
    MsgBox "Error From: " & strModuleName & vbNewLine & _
           "Occur In: " & strOccurIn & vbNewLine & _
           "Error Number: " & objError.Number & vbNewLine & _
           "Description: " & objError.Description, vbCritical, "Application Error"
    'Save the error log (The save error log will be display later on in the program)
    Open App.Path & "\Error.log" For Append As #1
        Print #1, Format(Date, "MMM-dd-yyyy") & "~~~~~" & Time & "~~~~~" & objError.Number & "~~~~~" & objError.Description & "~~~~~" & strModuleName & "~~~~~" & strOccurIn
    Close #1
End Sub


Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Integer, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String, Optional srcFieldForToolTip As String, Optional BooleanField As String, Optional iconPathField As String, Optional ilIconLarge As ImageList, Optional ilIconSmall As ImageList, Optional picTemp As PictureBox)
    Dim tempItem As ListItem
    Dim i As Byte
    On Error Resume Next
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.MoveFirst
    Do While Not sRecordSource.EOF
        If with_num = True Then
            If sNumIco = -1 Then
                Set tempItem = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, ExtractIcon(sRecordSource.Fields(iconPathField), ilIconLarge, picTemp, 32), ExtractIcon(sRecordSource.Fields(iconPathField), ilIconSmall, picTemp, 16))
            Else
                Set tempItem = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
            End If
        Else
            If sNumIco = -1 Then
                Set tempItem = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), ExtractIcon(sRecordSource.Fields(iconPathField), ilIconLarge, picTemp, 32), ExtractIcon(sRecordSource.Fields(iconPathField), ilIconSmall, picTemp, 16))
            Else
               Set tempItem = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
            End If
        End If
            If srcHiddenField <> "" Then tempItem.Tag = sRecordSource.Fields(srcHiddenField)
            If srcFieldForToolTip <> "" Then tempItem.ToolTipText = sRecordSource.Fields(srcFieldForToolTip)

            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                        If sRecordSource.Fields(CInt(i) - 1).Type = adDouble Then
                            tempItem.SubItems(i) = FormatRecord(sRecordSource.Fields(CInt(i) - 1), , BooleanField)
                        Else
                            tempItem.SubItems(i) = "" & FormatRecord(sRecordSource.Fields(CInt(i) - 1), , BooleanField)
                        End If
                    Else
                        If sRecordSource.Fields(CInt(i)).Type = adDouble Then
                            tempItem.SubItems(i) = FormatRecord(sRecordSource.Fields(CInt(i)), , BooleanField)
                        Else
                            tempItem.SubItems(i) = "" & FormatRecord(sRecordSource.Fields(CInt(i)), , BooleanField)
                        End If
                    End If
                Else
                    tempItem.SubItems(i) = "" & FormatRecord(sRecordSource.Fields(CInt(i) + 1), , BooleanField)
                End If
            Next i
        sRecordSource.MoveNext
    Loop
    i = 0
    Set tempItem = Nothing
End Sub

Public Sub PrompAccessDenied()
    MsgBox "You have no permission to do that task.Please contact your administrator.", vbCritical, "Access Denied"
End Sub

Public Sub PrompAccessDeniedForFolder()
    MsgBox "You have no permission to do that task.Please check your folder permission or contact your administrator.", vbCritical, "Access Denied"
End Sub

