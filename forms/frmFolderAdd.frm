VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFolderAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Folder"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolderAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   150
      ScaleHeight     =   390
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   1200
      Width           =   4515
      Begin VB.TextBox txtName 
         DataField       =   "FolderName"
         Height          =   315
         Left            =   1425
         TabIndex        =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Name:"
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3570
      TabIndex        =   2
      Top             =   4830
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2355
      TabIndex        =   1
      Top             =   4830
      Width           =   1140
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   3
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   225
      TabIndex        =   5
      Top             =   6300
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3990
      Left            =   75
      TabIndex        =   6
      Top             =   750
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   7038
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "New Folder"
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
      TabIndex        =   4
      Top             =   150
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmFolderAdd.frx":038A
      Top             =   150
      Width           =   360
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
Attribute VB_Name = "frmFolderAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lPK As Long
Public lParentPK As Long
Public strDirName As String

Dim bAddState As Boolean
Dim rsRecord As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(txtName) Then Exit Sub
    If frmFileManager.IsDirectoryExist(txtName.Text) = True Then Exit Sub
    
    ShowWaiting Me, WaitingDisplay
    
    txtName.Text = RemoveInvalidChar(txtName.Text, "'")
    SyncronizeRecordsetBinding rsRecord, Me
    
    rsRecord.Fields("ParentFolderID") = lParentPK
    rsRecord.Fields("DirectoryName") = strDirName
    rsRecord.Fields("DateCreated") = Now
    rsRecord.Fields("CreatedBy") = AppCurrentUser.CompleteName
    
    SaveRecord "SELECT [FolderName],[ParentFolderID],[DirectoryName],[DateCreated],[CreatedBy]" & _
           " FROM tbl_Folders WHERE [ID]=" & lPK, rsRecord, , Not bAddState, "tbl_Folders"
    
    If bAddState = True Then
        Set LastRecordsetA = GetRecords("SELECT * FROM tbl_Folders WHERE [ID]=" & LAST_GENERATED_IDENTITY)
    Else
        Set LastRecordsetA = rsRecord
    End If
    HideWaiting Me, WaitingDisplay
    Unload Me
    
End Sub

Private Sub Form_Load()
    If lPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Set rsRecord = GetRecords("SELECT * FROM tbl_Folders WHERE [ID]=" & lPK)
    If rsRecord Is Nothing Then
        MsgBox "Unable to read the data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        BindControlToRS Me, rsRecord
    End If
    
    If bAddState = True Then
        rsRecord.AddNew
    Else
    End If
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFolderAdd = Nothing
End Sub


