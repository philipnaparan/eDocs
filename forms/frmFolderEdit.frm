VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFolderEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Folder"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolderEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2325
      TabIndex        =   5
      Top             =   4830
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3570
      TabIndex        =   12
      Top             =   4830
      Width           =   1140
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   15
      Top             =   600
      Width           =   6840
      _extentx        =   12065
      _extenty        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   75
      TabIndex        =   16
      Top             =   5625
      Width           =   2640
      _extentx        =   4921
      _extenty        =   1349
   End
   Begin VB.PictureBox pnlGeneral 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   150
      ScaleHeight     =   3390
      ScaleWidth      =   4515
      TabIndex        =   13
      Top             =   1200
      Width           =   4515
      Begin VB.TextBox txtDateCreatedBy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   3
         Top             =   1350
         Width           =   3015
      End
      Begin VB.TextBox txtLastModifiedBy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   4
         Top             =   1725
         Width           =   3015
      End
      Begin VB.TextBox txtLastModified 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   2
         Top             =   975
         Width           =   3015
      End
      Begin VB.TextBox txtDateCreated 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1425
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtFolderName 
         DataField       =   "FolderName"
         Height          =   315
         Left            =   1425
         TabIndex        =   0
         Top             =   0
         Width           =   3015
      End
      Begin eDocs.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   75
         TabIndex        =   19
         Top             =   450
         Width           =   4365
         _extentx        =   7699
         _extenty        =   53
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Created By:"
         Height          =   315
         Index           =   15
         Left            =   75
         TabIndex        =   23
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified By:"
         Height          =   315
         Index           =   14
         Left            =   75
         TabIndex        =   22
         Top             =   1725
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified:"
         Height          =   315
         Index           =   13
         Left            =   75
         TabIndex        =   21
         Top             =   975
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Created:"
         Height          =   315
         Index           =   12
         Left            =   75
         TabIndex        =   20
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Name:"
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1365
      End
   End
   Begin MSComctlLib.TabStrip tabFileInfo 
      Height          =   3990
      Left            =   75
      TabIndex        =   17
      Top             =   750
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   7038
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "general"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Restrictions"
            Key             =   "restrict"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1350
      Top             =   75
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
            Picture         =   "frmFolderEdit.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pnlRestrict 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   150
      ScaleHeight     =   3465
      ScaleWidth      =   4515
      TabIndex        =   11
      Top             =   1125
      Width           =   4515
      Begin VB.CommandButton cmdDetailRefresh 
         Caption         =   "&Refresh"
         Height          =   390
         Left            =   3450
         TabIndex        =   10
         Top             =   2925
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailDelete 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   2325
         TabIndex        =   9
         Top             =   2925
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailEdit 
         Caption         =   "&Edit"
         Height          =   390
         Left            =   1200
         TabIndex        =   8
         Top             =   2925
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailAdd 
         Caption         =   "&New"
         Height          =   390
         Left            =   75
         TabIndex        =   7
         Top             =   2925
         Width           =   1005
      End
      Begin MSComctlLib.ListView lstvRestrictions 
         Height          =   2715
         Left            =   75
         TabIndex        =   6
         Top             =   150
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   4789
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
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group Name"
            Object.Width           =   7232
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmFolderEdit.frx":0724
      Top             =   150
      Width           =   360
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Folder"
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
      Left            =   600
      TabIndex        =   18
      Top             =   150
      Width           =   3690
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -150
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmFolderEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lFolderId As Long

Dim rsFolder As recordset
Dim strOldFolderName As String

Private Sub SaveFile()
    
    If Not rsFolder Is Nothing Then
        
        ShowWaiting Me, WaitingDisplay
        
        With rsFolder
            .Fields("LastModified") = Now
            .Fields("LastModifiedBy") = AppCurrentUser.CompleteName
            .Fields("FolderName") = txtFolderName.Text
            '.Update
        End With
        
        If SaveRecord("", rsFolder, , True, "tbl_Folders") = 1 Then
            frmFileManager.RenameSelNode txtFolderName.Text, lFolderId
        End If
    
        Set rsFolder = Nothing
        HideWaiting Me, WaitingDisplay
        
        Unload Me
    End If
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDetailAdd_Click()
    
    frmFolderRestrictionsAddEdit.lRecordPK = 0
    frmFolderRestrictionsAddEdit.lFolderId = lFolderId
    frmFolderRestrictionsAddEdit.Show vbModal
    
    LastGenericText = "HaveChanges"
    
    LoadRestrictionRecords
End Sub

Private Sub LoadRestrictionRecords()

    ShowWaiting Me, WaitingDisplay

    lstvRestrictions.ListItems.Clear

    Dim rsRestrict As recordset
    Set rsRestrict = GetRecords("SELECT [GroupName],[ID],[FolderId] FROM vw_FolderRestrictions WHERE [FolderId]=" & lFolderId & " ORDER BY [GroupName] ASC")

    If Not rsRestrict Is Nothing Then
        If rsRestrict.RecordCount > 0 Then
            DisableEditing True

            FillListView lstvRestrictions, rsRestrict, 1, 1, False, True, "ID", "GroupName"
        Else
            DisableEditing
        End If
    End If

    Set rsRestrict = Nothing

    HideWaiting Me, WaitingDisplay

End Sub

Private Sub DisableEditing(Optional EnableEditing As Boolean)
    cmdDetailEdit.Enabled = EnableEditing
    cmdDetailDelete.Enabled = EnableEditing
End Sub

Private Sub cmdDetailDelete_Click()
    If MsgBox("Are you sure you want to delete the restriction in the group named '" & lstvRestrictions.SelectedItem.Text & "'?", vbCritical + vbYesNo, "Confirm Restriction Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_FolderRestrictions WHERE [ID]=" & lstvRestrictions.SelectedItem.Tag & "") = True Then
            lstvRestrictions.ListItems.Remove lstvRestrictions.SelectedItem.Index
            lstvRestrictions.SelectedItem.Selected = True
            LastGenericText = "HaveChanges"
        End If
        HideWaiting Me, WaitingDisplay
    End If
End Sub

Private Sub lstvRestrictions_DblClick()
    cmdDetailEdit_Click
End Sub

Private Sub cmdDetailEdit_Click()
    
    frmFolderRestrictionsAddEdit.lRecordPK = Val(lstvRestrictions.SelectedItem.Tag)
    frmFolderRestrictionsAddEdit.lFolderId = lFolderId
    frmFolderRestrictionsAddEdit.Show vbModal
    LastGenericText = "HaveChanges"

    LoadRestrictionRecords
End Sub

Private Sub cmdDetailRefresh_Click()
  LoadRestrictionRecords
End Sub


Private Sub cmdSave_Click()
    If IsControlEmpty(txtFolderName) Then Exit Sub
    If strOldFolderName <> txtFolderName.Text Then
        ShowWaiting Me, WaitingDisplay
        If frmFileManager.IsDirectoryExist(txtFolderName.Text, True) = True Then HideWaiting Me, WaitingDisplay: Exit Sub
        HideWaiting Me, WaitingDisplay
    End If
    
    SaveFile
End Sub



Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    pnlGeneral.BackColor = &H8000000F
    pnlRestrict.BackColor = &H8000000F
    
    pnlGeneral.ZOrder
    
    
    DoEvents

    ShowWaiting Me, WaitingDisplay
    
    Set rsFolder = GetRecords("SELECT [FolderName],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy] FROM tbl_Folders WHERE [ID]=" & lFolderId & "")

            
    If Not rsFolder Is Nothing Then
        If rsFolder.RecordCount > 0 Then
            
            On Error Resume Next
            
            If IsNull(rsFolder.Fields("FolderName")) = False Then txtFolderName.Text = rsFolder.Fields("FolderName")

            If IsNull(rsFolder.Fields("DateCreated")) = False Then txtDateCreated.Text = rsFolder.Fields("DateCreated")
            If IsNull(rsFolder.Fields("LastModified")) = False Then txtLastModified.Text = rsFolder.Fields("LastModified")
            If IsNull(rsFolder.Fields("CreatedBy")) = False Then txtDateCreatedBy.Text = rsFolder.Fields("CreatedBy")
            If IsNull(rsFolder.Fields("LastModifiedBy")) = False Then txtLastModifiedBy.Text = rsFolder.Fields("LastModifiedBy")
            
        End If
        strOldFolderName = txtFolderName.Text
        
        LoadRestrictionRecords
        
    End If
    
    HideWaiting Me, WaitingDisplay
       
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmFolderEdit = Nothing
End Sub


Private Sub tabFileInfo_Click()
    pnlGeneral.Visible = False
    pnlRestrict.Visible = False
    
    Select Case tabFileInfo.SelectedItem.key
        Case "general"
            pnlGeneral.Visible = True
            pnlGeneral.ZOrder
        Case "restrict"
            pnlRestrict.Visible = True
            pnlRestrict.ZOrder
    End Select
End Sub


