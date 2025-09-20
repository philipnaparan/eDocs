Attribute VB_Name = "modControls"

Option Explicit

Public Sub UpdateListView(ByRef ListViewControl As ListView, _
                            ByRef RecordSource As recordset, _
                            ByVal IconIndex As Integer, _
                            Optional HiddenField As String, _
                            Optional ToolTipField As String, _
                            Optional BooleanField As String, _
                            Optional iconPathField As String, _
                            Optional LargeIconImageList As ImageList, _
                            Optional SmallIconImageList As ImageList, _
                            Optional TempPic As PictureBox, _
                            Optional IsInsert As Boolean)
                            
    On Error Resume Next
    
    Dim tempItem As ListItem
    Dim i As Long
    
    If RecordSource.RecordCount < 1 Then Exit Sub
    
    If IsInsert = True Then
        If IconIndex = -1 Then
            Set tempItem = ListViewControl.ListItems.Add(, , "" & RecordSource.Fields(ListViewControl.ColumnHeaders(1).key), ExtractIcon(RecordSource.Fields(iconPathField), LargeIconImageList, TempPic, 32), ExtractIcon(RecordSource.Fields(iconPathField), SmallIconImageList, TempPic, 16))
        Else
            Set tempItem = ListViewControl.ListItems.Add(, , "" & RecordSource.Fields(ListViewControl.ColumnHeaders(1).key), IconIndex, IconIndex)
        End If
        
        tempItem.Selected = True
        tempItem.EnsureVisible
    Else
        If ListViewControl.SelectedItem Is Nothing Then Exit Sub
        Set tempItem = ListViewControl.SelectedItem
        
        tempItem.Text = "" & RecordSource.Fields(ListViewControl.ColumnHeaders(1).key)
    End If
    
    If HiddenField <> "" Then tempItem.Tag = RecordSource.Fields(HiddenField)
    If ToolTipField <> "" Then tempItem.ToolTipText = RecordSource.Fields(ToolTipField)

    If ListViewControl.ColumnHeaders.Count > 1 Then
        For i = 2 To ListViewControl.ColumnHeaders.Count
            If RecordSource.Fields(ListViewControl.ColumnHeaders(i).key).Type = adDouble Then
                tempItem.SubItems(i - 1) = FormatRecord(RecordSource.Fields(ListViewControl.ColumnHeaders(i).key), , BooleanField)
            Else
                tempItem.SubItems(i - 1) = "" & FormatRecord(RecordSource.Fields(ListViewControl.ColumnHeaders(i).key), , BooleanField)
            End If
        Next i
        i = 0
    End If
    
    Set tempItem = Nothing
End Sub

Public Sub ShowWaiting(ByRef Form As Form, ByRef WaitingDisplay As ctrlWaiting)
    
    Form.WaitingDisplay.ZOrder
    Form.WaitingDisplay.Visible = True
    Form.WaitingDisplay.Left = (Form.Width - Form.WaitingDisplay.Width) / 2
    Form.WaitingDisplay.Top = (Form.Height - Form.WaitingDisplay.Height) / 2
    Form.WaitingDisplay.StartAnim

    Form.MousePointer = vbHourglass
    
    Wait 1
    
    Form.Enabled = False
    
    
End Sub

Public Sub HideWaiting(ByRef Form As Form, ByRef WaitingDisplay As ctrlWaiting)

    'Form.WaitingDisplay.Visible = False
    Form.WaitingDisplay.Terminate
    Form.MousePointer = vbDefault
    Form.Enabled = True
End Sub


Public Function IsControlEmpty(ByRef ObjControl) As Boolean
    On Error Resume Next
    If ObjControl.Text = "" Then
        Beep
        ObjControl.SetFocus
        IsControlEmpty = True
    End If
End Function

Public Function IsControlTagEmpty(ByRef ObjControl) As Boolean
    If ObjControl.Tag = "" Then
        Beep
        ObjControl.SetFocus
        IsControlTagEmpty = True
    End If
End Function

Public Sub HighLightText(ByRef sText)
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub


Public Sub FillComboBox(ByVal recordset As recordset, ByRef combo As ComboBox, ByVal FieldName As String)
    combo.Clear
    If recordset.RecordCount > 0 Then
        recordset.MoveFirst
        Do While Not recordset.EOF
            combo.AddItem recordset.Fields(FieldName)
            recordset.MoveNext
        Loop
    End If
    
End Sub
