VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUpdateFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update File Info"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdateFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2475
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEditAttachment 
      Caption         =   "&Edit Attachment"
      Height          =   390
      Left            =   75
      TabIndex        =   26
      Top             =   6300
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2550
      TabIndex        =   28
      Top             =   6300
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   390
      Left            =   3900
      TabIndex        =   27
      Top             =   6300
      Width           =   1455
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   2400
      TabIndex        =   52
      Top             =   7425
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.PictureBox pnlMetaData 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   150
      ScaleHeight     =   4815
      ScaleWidth      =   5190
      TabIndex        =   48
      Top             =   1125
      Width           =   5190
      Begin VB.CommandButton cmdDetailRefresh 
         Caption         =   "&Refresh"
         Height          =   390
         Left            =   3450
         TabIndex        =   17
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailDelete 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   2325
         TabIndex        =   16
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailEdit 
         Caption         =   "&Edit"
         Height          =   390
         Left            =   1200
         TabIndex        =   15
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailAdd 
         Caption         =   "&New"
         Height          =   390
         Left            =   75
         TabIndex        =   14
         Top             =   4350
         Width           =   1005
      End
      Begin MSComctlLib.ListView lstvProperty 
         Height          =   4140
         Left            =   75
         TabIndex        =   13
         Top             =   150
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   7303
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Property"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   4128
         EndProperty
      End
   End
   Begin VB.PictureBox pnlHistory 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   150
      ScaleHeight     =   3615
      ScaleWidth      =   5115
      TabIndex        =   42
      Top             =   1125
      Width           =   5115
      Begin VB.TextBox txtDateCheckOut 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   20
         Top             =   750
         Width           =   3690
      End
      Begin VB.TextBox txtCheckOutBy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   19
         Top             =   375
         Width           =   3690
      End
      Begin VB.TextBox txtFileSize 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   18
         Top             =   0
         Width           =   3690
      End
      Begin VB.TextBox txtDateCreated 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   21
         Top             =   1275
         Width           =   3690
      End
      Begin VB.TextBox txtLastModified 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   22
         Top             =   1650
         Width           =   3690
      End
      Begin VB.TextBox txtLastModifiedBy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   24
         Top             =   2400
         Width           =   3690
      End
      Begin VB.TextBox txtDateCreatedBy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         TabIndex        =   23
         Top             =   2025
         Width           =   3690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Check Out:"
         Height          =   315
         Index           =   18
         Left            =   0
         TabIndex        =   54
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Out By:"
         Height          =   315
         Index           =   17
         Left            =   0
         TabIndex        =   53
         Top             =   375
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File Size:"
         Height          =   315
         Index           =   11
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Created:"
         Height          =   315
         Index           =   12
         Left            =   0
         TabIndex        =   46
         Top             =   1275
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified:"
         Height          =   315
         Index           =   13
         Left            =   0
         TabIndex        =   45
         Top             =   1650
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified By:"
         Height          =   315
         Index           =   14
         Left            =   0
         TabIndex        =   44
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Created By:"
         Height          =   315
         Index           =   15
         Left            =   0
         TabIndex        =   43
         Top             =   2025
         Width           =   1245
      End
   End
   Begin VB.PictureBox pnlGeneral 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   150
      ScaleHeight     =   4815
      ScaleWidth      =   5115
      TabIndex        =   31
      Top             =   1125
      Width           =   5115
      Begin VB.ComboBox cbxStatus 
         Height          =   315
         Left            =   1500
         TabIndex        =   25
         Top             =   4425
         Width           =   3540
      End
      Begin VB.CheckBox ckConfidential 
         Caption         =   "Mark As Confidential"
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Top             =   4125
         Width           =   1860
      End
      Begin VB.TextBox txtPhysicalLoc 
         Height          =   315
         Left            =   1500
         TabIndex        =   9
         Top             =   3750
         Width           =   3540
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   375
         Width           =   3540
      End
      Begin VB.ComboBox cbxDocType 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   3375
         Width           =   3540
      End
      Begin VB.TextBox txtIndex 
         Height          =   315
         Left            =   1500
         TabIndex        =   6
         Top             =   2625
         Width           =   3540
      End
      Begin VB.TextBox txtNo 
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Top             =   2250
         Width           =   3540
      End
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   1875
         Width           =   3540
      End
      Begin VB.TextBox txtDesc 
         Height          =   1065
         Left            =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   750
         Width           =   3540
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   0
         Width           =   3540
      End
      Begin MSComCtl2.DTPicker dtpDocDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Top             =   3000
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   75825153
         CurrentDate     =   39436
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   315
         Index           =   16
         Left            =   0
         TabIndex        =   41
         Top             =   4425
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   39
         Top             =   675
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Original Author:"
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Top             =   1875
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. No:"
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   37
         Top             =   2250
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Index:"
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   36
         Top             =   2625
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Date:"
         Height          =   315
         Index           =   5
         Left            =   0
         TabIndex        =   35
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Type:"
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   34
         Top             =   3375
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Physical Location:"
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   33
         Top             =   3750
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   315
         Index           =   10
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   375
         Width           =   1395
      End
   End
   Begin MSComctlLib.TabStrip tabFileInfo 
      Height          =   5490
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   9684
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "general"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alerts"
            Key             =   "alert"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extra Properties"
            Key             =   "metadata"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Others"
            Key             =   "others"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pnlAlert 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   4590
      Left            =   150
      ScaleHeight     =   4590
      ScaleWidth      =   5190
      TabIndex        =   49
      Top             =   1125
      Width           =   5190
      Begin MSComCtl2.DTPicker dtpDocExp 
         Height          =   315
         Left            =   1275
         TabIndex        =   11
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   75825153
         CurrentDate     =   39436
      End
      Begin VB.TextBox txtEmpNote 
         Height          =   3690
         Left            =   1275
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   375
         Width           =   3840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Date:"
         Height          =   315
         Index           =   8
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Note:"
         Height          =   315
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   375
         Width           =   1245
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmUpdateFile.frx":0724
      Top             =   150
      Width           =   360
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Update File info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   675
      TabIndex        =   29
      Top             =   150
      Width           =   5865
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmUpdateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lFileId As Long
Public lFolderId As Long
Public IsEditMode As Boolean

Dim strFileExt As String
Dim strDocFileLocation As String


Dim rsDocument As recordset

Private Sub SaveFile()
    
    
    
    If Not rsDocument Is Nothing Then
        
        ShowWaiting Me, WaitingDisplay
        
        With rsDocument
                       
            If IsEditMode = True Then
                SaveFileToDB strDocFileLocation, rsDocument, "FileObj"
                .Fields("FileSize") = GetFileSizeString(strDocFileLocation)
                
                .Fields("LastModified") = Now
                .Fields("LastModifiedBy") = AppCurrentUser.CompleteName
            Else
                .Fields("FolderID") = lFolderId
                
                .Fields("LastModified") = Null
                .Fields("LastModifiedBy") = Null
            End If
            .Fields("FileName") = txtFileName.Text & "." & strFileExt
            
            .Fields("Title") = txtTitle.Text
            .Fields("Description") = txtDesc.Text
            .Fields("OriginalAuthor") = txtAuthor.Text
            .Fields("DocumentNo") = txtNo.Text
            .Fields("DocumentIndex") = txtIndex.Text
            If IsNull(dtpDocDate.Value) Then
                .Fields("DocumentDate") = dtpDocDate.Value
            Else
                .Fields("DocumentDate") = CDate(dtpDocDate.Value)
            End If
            .Fields("DocumentType") = cbxDocType.Text
            .Fields("PhysicalLocation") = txtPhysicalLoc.Text
            If IsNull(dtpDocExp.Value) Then
                .Fields("AlertDate") = dtpDocExp.Value
            Else
                .Fields("AlertDate") = CDate(dtpDocExp.Value)
            End If
            .Fields("AlertNote") = txtEmpNote.Text
            .Fields("IsConfidential") = ckConfidential.Value
            .Fields("Status") = cbxStatus.Text
            
            '.Update
            
        End With
        
        If IsEditMode = True Then
            If SaveRecord("SELECT [FileName],[Title],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[FileSize],[FileSizeNo],[AlertDate],[AlertNote],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy],[IsConfidential],[Description],[FileObj],[FolderID],[Status] FROM tbl_Files " & _
                          "WHERE [ID]=" & lFileId & "", rsDocument, , True, "tbl_Files") = 1 Then
                
                
                frmFileManager.docViewer.UpdateList rsDocument
                
            End If
        Else
            If SaveRecord("SELECT [FileName],[Title],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[FileSize],[AlertDate],[AlertNote],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy],[IsConfidential],[Description],[FolderID],[Status] FROM tbl_Files " & _
                          "WHERE [ID]=" & lFileId & "", rsDocument, , True, "tbl_Files") = 1 Then
                
                
                frmFileManager.docViewer.UpdateList rsDocument
                
            End If
        End If
    
        Set rsDocument = Nothing
        HideWaiting Me, WaitingDisplay
        
        Unload Me
    End If
    
    
   
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDetailAdd_Click()
    
    frmFilePropertyAddEdit.lRecordPK = 0
    frmFilePropertyAddEdit.lFileId = lFileId
    frmFilePropertyAddEdit.Show vbModal
    
    LoadPropertyRecords
End Sub

Private Sub LoadPropertyRecords()

    ShowWaiting Me, WaitingDisplay
    
    lstvProperty.ListItems.Clear
    
    Dim rsProperty As recordset
    Set rsProperty = GetRecords("SELECT [PropertyName],[PropertyValue],[ID],[FileId] FROM tbl_FileProperty WHERE [FileId]=" & lFileId & " ORDER BY [PropertyName] ASC")
    
    If Not rsProperty Is Nothing Then
        If rsProperty.RecordCount > 0 Then
            DisableEditing True
            
            FillListView lstvProperty, rsProperty, 2, 1, False, True, "ID", "PropertyName"
        Else
            DisableEditing
        End If
    End If
    
    Set rsProperty = Nothing
    
    HideWaiting Me, WaitingDisplay
    
End Sub


Private Sub DisableEditing(Optional EnableEditing As Boolean)
    cmdDetailEdit.Enabled = EnableEditing
    cmdDetailDelete.Enabled = EnableEditing
End Sub

Private Sub cmdDetailDelete_Click()
    If MsgBox("Are you sure you want to delete the property named '" & lstvProperty.SelectedItem.Text & "'?", vbCritical + vbYesNo, "Confirm Property Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_FileProperty WHERE [ID]=" & lstvProperty.SelectedItem.Tag & "") = True Then
            lstvProperty.ListItems.Remove lstvProperty.SelectedItem.Index
            lstvProperty.SelectedItem.Selected = True
        End If
        HideWaiting Me, WaitingDisplay
    End If
End Sub

Private Sub lstvProperty_DblClick()
    cmdDetailEdit_Click
End Sub

Private Sub cmdDetailEdit_Click()
    
    frmFilePropertyAddEdit.lRecordPK = Val(lstvProperty.SelectedItem.Tag)
    frmFilePropertyAddEdit.lFileId = lFileId
    frmFilePropertyAddEdit.Show vbModal
    
    LoadPropertyRecords
End Sub

Private Sub cmdDetailRefresh_Click()
  LoadPropertyRecords
End Sub

Private Sub cmdEditAttachment_Click()
    frmFileLuncher.fileName = strDocFileLocation
    frmFileLuncher.Show vbModal
    
    rsDocument.Fields("FileSize") = GetFileSizeString(strDocFileLocation)
    rsDocument.Fields("FileSizeNo") = GetFileSize(strDocFileLocation)
    
End Sub

Private Sub cmdSave_Click()
    SaveFile
End Sub



Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    pnlGeneral.BackColor = &H8000000F
    pnlAlert.BackColor = &H8000000F
    pnlMetaData.BackColor = &H8000000F
    pnlHistory.BackColor = &H8000000F
    
    pnlGeneral.ZOrder
    
    dtpDocDate.Value = Date
    dtpDocExp.Value = Date
    
    dtpDocDate.Value = Null
    dtpDocExp.Value = Null
    DoEvents

    ShowWaiting Me, WaitingDisplay
    
    If IsEditMode = True Then
        Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & lFileId & "")
        cmdEditAttachment.Visible = True
    Else
        Set rsDocument = GetRecords("SELECT * FROM vw_FileInfoOnly WHERE [ID]=" & lFileId & "")
        cmdEditAttachment.Visible = False
    End If
            
    If Not rsDocument Is Nothing Then
        If rsDocument.RecordCount > 0 Then
            
            If IsEditMode = True Then
                strDocFileLocation = App.Path & "\temp\" & GetFileNameWithOutExt(rsDocument.Fields("FileName")) & "_" & GenerateFileName("." & GetFileExt(rsDocument.Fields("FileName")))
            
                If DownloadFileFromDB(strDocFileLocation, rsDocument, "FileObj") = False Then
                    MsgBox "Error occur while downloading the document.", vbCritical
                    Unload Me
                    Exit Sub
                End If
            End If
            
            Dim rsDropList As recordset
   
            Set rsDropList = GetRecords("SELECT * FROM tbl_DocType ORDER BY TypeName ASC")
            If Not rsDropList Is Nothing Then
                 FillComboBox rsDropList, cbxDocType, "TypeName"
                 Set rsDropList = Nothing
            End If
            
            Set rsDropList = GetRecords("SELECT * FROM tbl_FileStatus ORDER BY StatusName ASC")
            If Not rsDropList Is Nothing Then
                 FillComboBox rsDropList, cbxStatus, "StatusName"
                 Set rsDropList = Nothing
            End If
            
            On Error Resume Next
            
            HideWaiting Me, WaitingDisplay
    
'            If IsEditMode = True Then
'                frmFileLuncher.fileName = strDocFileLocation
'                frmFileLuncher.Show vbModal
'
'                rsDocument.Fields("FileSize") = GetFileSizeString(strDocFileLocation)
'
'            End If
            
            
            ShowWaiting Me, WaitingDisplay
            
            If IsNull(rsDocument.Fields("Title")) = False Then txtTitle.Text = rsDocument.Fields("Title")
            If IsNull(rsDocument.Fields("Description")) = False Then txtDesc.Text = rsDocument.Fields("Description")
            If IsNull(rsDocument.Fields("OriginalAuthor")) = False Then txtAuthor.Text = rsDocument.Fields("OriginalAuthor")
            If IsNull(rsDocument.Fields("DocumentNo")) = False Then txtNo.Text = rsDocument.Fields("DocumentNo")
            If IsNull(rsDocument.Fields("DocumentIndex")) = False Then txtIndex.Text = rsDocument.Fields("DocumentIndex")
            If IsNull(rsDocument.Fields("DocumentDate")) = False Then dtpDocDate.Value = rsDocument.Fields("DocumentDate")
            If IsNull(rsDocument.Fields("DocumentType")) = False Then cbxDocType.Text = rsDocument.Fields("DocumentType")
            If IsNull(rsDocument.Fields("PhysicalLocation")) = False Then txtPhysicalLoc.Text = rsDocument.Fields("PhysicalLocation")
            If IsNull(rsDocument.Fields("AlertDate")) = False Then dtpDocExp.Value = rsDocument.Fields("AlertDate")
            If IsNull(rsDocument.Fields("AlertNote")) = False Then txtEmpNote.Text = rsDocument.Fields("AlertNote")
            If IsNull(rsDocument.Fields("IsConfidential")) = False Then ckConfidential.Value = Val(rsDocument.Fields("IsConfidential"))
            If IsNull(rsDocument.Fields("Status")) = False Then cbxStatus.Text = rsDocument.Fields("Status")

            If IsNull(rsDocument.Fields("FileSize")) = False Then txtFileSize.Text = rsDocument.Fields("FileSize")
            If IsNull(rsDocument.Fields("CheckOutBy")) = False Then txtCheckOutBy.Text = rsDocument.Fields("CheckOutBy")
            If IsNull(rsDocument.Fields("DateCheckOut")) = False Then txtDateCheckOut.Text = rsDocument.Fields("DateCheckOut")
            
            If IsNull(rsDocument.Fields("DateCreated")) = False Then txtDateCreated.Text = rsDocument.Fields("DateCreated")
            If IsNull(rsDocument.Fields("LastModified")) = False Then txtLastModified.Text = rsDocument.Fields("LastModified")
            If IsNull(rsDocument.Fields("CreatedBy")) = False Then txtDateCreatedBy.Text = rsDocument.Fields("CreatedBy")
            If IsNull(rsDocument.Fields("LastModifiedBy")) = False Then txtLastModifiedBy.Text = rsDocument.Fields("LastModifiedBy")
 
            If IsNull(rsDocument.Fields("FileName")) = False Then strFileExt = GetFileExt(rsDocument.Fields("FileName"))
            If IsNull(rsDocument.Fields("FileName")) = False Then txtFileName.Text = GetFileNameWithOutExt(rsDocument.Fields("FileName"))

            If IsEditMode = False Then
                txtTitle.Text = txtFileName.Text
                txtAuthor.Text = txtDateCreatedBy.Text
            Else
                If txtDateCheckOut.Text <> "" Then
                    MsgBox "This file is currently check-out.", vbInformation
                    cmdEditAttachment.Enabled = False
                    cmdSave.Enabled = False
                End If
            End If
        End If
        
        LoadPropertyRecords
        
    End If
    
    HideWaiting Me, WaitingDisplay
       
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Shell App.Path & "\temp\TempEraser.exe"
    
    Set frmUpdateFile = Nothing
End Sub


Private Sub tabFileInfo_Click()
    pnlGeneral.Visible = False
    pnlAlert.Visible = False
    pnlMetaData.Visible = False
    pnlHistory.Visible = False
    
    Select Case tabFileInfo.SelectedItem.key
        Case "general"
            pnlGeneral.Visible = True
            pnlGeneral.ZOrder
        Case "alert"
            pnlAlert.Visible = True
            pnlAlert.ZOrder
        Case "metadata"
            pnlMetaData.Visible = True
            pnlMetaData.ZOrder
        Case "others"
            pnlHistory.Visible = True
            pnlHistory.ZOrder
    End Select
End Sub
