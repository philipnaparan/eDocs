VERSION 5.00
Object = "{0DB2A541-9FC9-41FE-8869-62AF866AA3F8}#1.0#0"; "OA.ocx"
Begin VB.Form frmOfficeEditor 
   Caption         =   "MS Office Editor"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOfficeEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin OALib.OA WordEditor 
      Height          =   6315
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   10860
      _Version        =   65536
      _ExtentX        =   19156
      _ExtentY        =   11139
      _StockProps     =   0
   End
   Begin VB.PictureBox panelDesc 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   105
      ScaleHeight     =   1800
      ScaleWidth      =   10830
      TabIndex        =   7
      Top             =   7035
      Width           =   10830
      Begin VB.TextBox txtDesc 
         Height          =   1620
         Left            =   0
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         ToolTipText     =   "Please enter the description of this document here!"
         Top             =   210
         Width           =   10860
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "  Document Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2955
      End
   End
   Begin VB.CheckBox ckConfidential 
      Caption         =   "Mark As Confidential"
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Top             =   225
      Width           =   1860
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8055
      TabIndex        =   0
      Top             =   225
      Width           =   1230
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   3690
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaveAndClose 
      Caption         =   "&Save and Close"
      Height          =   315
      Left            =   9360
      TabIndex        =   2
      Top             =   225
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Note: Press Ctrl+P to print the document."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   675
      TabIndex        =   6
      Top             =   360
      Width           =   3210
   End
   Begin VB.Label Label1 
      Caption         =   "MS Office Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   675
      TabIndex        =   4
      Top             =   90
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmOfficeEditor.frx":038A
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmOfficeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lFileId As Long
Public lFolderID As Long
Public strDocName As String
Public strDocType As String

Dim strDocFileLocation As String
Dim bIsDocSaved As Boolean
Dim bCreateNew As Boolean

Dim rsDocument As recordset

Private Sub cmdPrint_Click()
    On Error Resume Next
    
    'WordEditor.SetFocus
    'SendKeys "{CTRL} + P"
End Sub

Private Sub cmdRename_Click()

    Dim strNewDocName As String
    Dim strOldDocName As String
    Dim strFileExt As String
    
    strFileExt = GetFileExt(strDocName)
    strNewDocName = GetNameFromFileName(strDocName)
    strOldDocName = strNewDocName
    
    strNewDocName = InputBox("Enter the new name of file named '" & strOldDocName & "':", "Rename File", strOldDocName)
    If strNewDocName = "" Or strNewDocName = strOldDocName Then Exit Sub
        
    strNewDocName = RemoveInvalidChar(strNewDocName, "'")
    strDocName = strNewDocName & "." & strFileExt
    
    'Update the display
    Me.Caption = "Edit Document - " & strDocName
    
End Sub

Private Sub cmdSaveAndClose_Click()

    If LCase(AppCurrentUser.UserType) = "viewer" Then PrompAccessDenied: Exit Sub
    If LCase(AppCurrentUser.UserType) = "viewer w/ confidential access" Then PrompAccessDenied: Exit Sub


    bIsDocSaved = False

    If bCreateNew = True Then
        CreateNewDoc
    Else
        updateDoc
    End If
    
    If bIsDocSaved = True Then Unload Me
    
End Sub

Private Sub CreateNewDoc()

    
    WordEditor.Save (strDocFileLocation)
        
    Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=-1")
    
    If Not rsDocument Is Nothing Then
        
        With rsDocument
            
            .AddNew
            
            .Fields("FileName") = strDocName
            
            SaveFileToDB strDocFileLocation, rsDocument, "FileObj"

            .Fields("Description") = txtDesc.Text
            .Fields("FolderID") = lFolderID
            .Fields("DateCreated") = Now
            .Fields("CreatedBy") = AppCurrentUser.CompleteName
            .Fields("IsConfidential") = ckConfidential.Value
            
            .Update
            
        End With
        
        If SaveRecord("SELECT [FileName],[Description],[FileObj],[FolderID],[DateCreated],[CreatedBy],[IsConfidential] FROM tbl_Files " & _
                      "WHERE [ID]=-1", rsDocument, , , "tbl_Files") = 1 Then
            
            
            frmFileManager.docViewer.UpdateListForNewDoc rsDocument
            bIsDocSaved = True
            
        End If
    
        Set rsDocument = Nothing
    End If
   
    
End Sub

Private Sub updateDoc()

    
    WordEditor.Save (strDocFileLocation)
        
    
    If Not rsDocument Is Nothing Then
        
        With rsDocument
                       
            .Fields("FileName") = strDocName
            
            SaveFileToDB strDocFileLocation, rsDocument, "FileObj"

            .Fields("Description") = txtDesc.Text
            .Fields("FolderID") = lFolderID
            .Fields("LastModified") = Now
            .Fields("LastModifiedBy") = AppCurrentUser.CompleteName
            .Fields("IsConfidential") = ckConfidential.Value
            
            .Update
            
        End With
        
        If SaveRecord("SELECT [FileName],[Description],[FileObj],[FolderID],[LastModified],[LastModifiedBy],[IsConfidential] FROM tbl_Files " & _
                      "WHERE [ID]=" & lFileId & "", rsDocument, , True, "tbl_Files") = 1 Then
            
            
            frmFileManager.docViewer.UpdateList rsDocument
            bIsDocSaved = True
            
        End If
    
        Set rsDocument = Nothing
    End If
   
    
End Sub


Private Sub Form_Load()
    DoEvents

    WordEditor.ShowToolbars True
    

    If lFileId = 0 Then
        bCreateNew = True
    Else
        bCreateNew = False
    End If

    If bCreateNew = True Then
        strDocName = strDocName & "." & strDocType
        Me.Caption = "New Document - " & strDocName
        cmdRename.Enabled = False
        
        strDocFileLocation = App.Path & "\temp\" & strDocName
        
        Select Case strDocType
            Case "doc"
                If WordEditor.CreateNew("Word.Document") = False Then
                    MsgBox "Error occur while creating new document.", vbCritical
                    Unload Me
                    Exit Sub
                End If
            Case "xls"
                If WordEditor.CreateNew("Excel.Sheet") = False Then
                    MsgBox "Error occur while creating new document.", vbCritical
                    Unload Me
                    Exit Sub
                End If
            Case "ppt"
                If WordEditor.CreateNew("PowerPoint.Show") = False Then
                    MsgBox "Error occur while creating new document.", vbCritical
                    Unload Me
                    Exit Sub
                End If
        End Select
        
    Else
               
        Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & lFileId & "")
            
        If Not rsDocument Is Nothing Then
            If rsDocument.RecordCount > 0 Then
            
                strDocName = rsDocument.Fields("FileName")
                strDocFileLocation = App.Path & "\temp\" & GenerateFileName("." & GetFileExt(strDocName))
            
                If DownloadFileFromDB(strDocFileLocation, rsDocument, "FileObj") = False Then
                    MsgBox "Error occur while downloading the document.", vbCritical
                    Unload Me
                    Exit Sub
                End If
                
                ckConfidential.Value = Val(rsDocument.Fields("IsConfidential"))
                txtDesc.Text = rsDocument.Fields("Description")
                
            End If
        End If
        
        If WordEditor.Open(strDocFileLocation) = False Then
            MsgBox "Error occur while reading the document.", vbCritical
            Unload Me
            Exit Sub
        Else
            Me.Caption = "Edit Document - " & strDocName
            cmdRename.Enabled = True
        End If

    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If bIsDocSaved = False Then
        Dim msgResult As VbMsgBoxResult
        msgResult = MsgBox("Do you want to save the changes in the file named '" & strDocName & "'?", vbExclamation + vbYesNoCancel)
        If msgResult = vbCancel Then
            Cancel = 1
        ElseIf msgResult = vbYes Then
            cmdSaveAndClose_Click
        End If
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    WordEditor.Width = Me.ScaleWidth - 180
    WordEditor.Height = Me.ScaleHeight - (WordEditor.Top + txtDesc.Height + 100)
    
    panelDesc.Top = WordEditor.Top + WordEditor.Height + 50
    panelDesc.Width = WordEditor.Width
    txtDesc.Width = WordEditor.Width
    
    cmdSaveAndClose.Left = (WordEditor.Left + WordEditor.Width) - cmdSaveAndClose.Width
'    cmdPrint.Left = cmdSaveAndClose.Left - (cmdPrint.Width + 50)
'    cmdRename.Left = cmdPrint.Left - (cmdRename.Width + 50)
'    ckConfidential.Left = cmdRename.Left - (ckConfidential.Width + 50)
    cmdRename.Left = cmdSaveAndClose.Left - (cmdPrint.Width + 50)
    ckConfidential.Left = cmdRename.Left - (ckConfidential.Width + 50)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell App.Path & "\temp\TempEraser.exe"
    
    Set frmOfficeEditor = Nothing
End Sub

