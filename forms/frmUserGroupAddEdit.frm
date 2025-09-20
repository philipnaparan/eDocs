VERSION 5.00
Begin VB.Form frmUserGroupAddEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserGroupAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check12 
      Caption         =   "Can Export"
      DataField       =   "CanExport"
      Height          =   240
      Left            =   1800
      TabIndex        =   7
      Top             =   3075
      Width           =   2190
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Can Import"
      DataField       =   "CanImport"
      Height          =   240
      Left            =   1800
      TabIndex        =   6
      Top             =   2775
      Width           =   2190
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Can Manage Templates"
      DataField       =   "CanManageTemplates"
      Height          =   240
      Left            =   1800
      TabIndex        =   11
      Top             =   4350
      Width           =   2190
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Can Add Folder"
      DataField       =   "CanAddFolder"
      Height          =   240
      Left            =   1800
      TabIndex        =   8
      Top             =   3450
      Width           =   2190
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Can Edit Folder"
      DataField       =   "CanEditFolder"
      Height          =   240
      Left            =   1800
      TabIndex        =   9
      Top             =   3750
      Width           =   2190
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Can Delete Folder"
      DataField       =   "CanDeleteFolder"
      Height          =   240
      Left            =   1800
      TabIndex        =   10
      Top             =   4050
      Width           =   2190
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Is System Administrator"
      DataField       =   "IsSystemAdministrator"
      Height          =   240
      Left            =   1800
      TabIndex        =   12
      Top             =   4725
      Width           =   2190
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Can Delete File"
      DataField       =   "CanDelete"
      Height          =   240
      Left            =   1800
      TabIndex        =   5
      Top             =   2475
      Width           =   2190
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Can Check Out File"
      DataField       =   "CanCheckOut"
      Height          =   240
      Left            =   1800
      TabIndex        =   4
      Top             =   2175
      Width           =   2190
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Can Edit File"
      DataField       =   "CanEdit"
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Top             =   1875
      Width           =   2190
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Can Add File"
      DataField       =   "CanAdd"
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   1575
      Width           =   2190
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Can View Confidential File"
      DataField       =   "CanViewConfidential"
      Height          =   240
      Left            =   1800
      TabIndex        =   1
      Top             =   1275
      Width           =   2190
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2430
      TabIndex        =   13
      Top             =   5400
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3675
      TabIndex        =   14
      Top             =   5400
      Width           =   1140
   End
   Begin VB.TextBox txtGroupName 
      DataField       =   "GroupName"
      Height          =   315
      Left            =   1785
      TabIndex        =   0
      Top             =   780
      Width           =   3015
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -225
      TabIndex        =   18
      Top             =   5250
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1125
      TabIndex        =   19
      Top             =   6150
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Folders:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   22
      Top             =   3450
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Others:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   21
      Top             =   4725
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Files:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   20
      Top             =   1275
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmUserGroupAddEdit.frx":038A
      Top             =   150
      Width           =   360
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   16
      Top             =   150
      Width           =   3690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   135
      TabIndex        =   15
      Top             =   780
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -75
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmUserGroupAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lGroupUserPK As Long

Dim bAddState As Boolean
Dim rsGroupUser As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(txtGroupName) Then Exit Sub
    
    ShowWaiting Me, WaitingDisplay
    
    SyncronizeRecordsetBinding rsGroupUser, Me
    SaveRecord "SELECT [GroupName]" & _
           ",[CanViewConfidential]" & _
           ",[CanAdd]" & _
           ",[CanEdit]" & _
           ",[CanCheckOut]" & _
           ",[CanDelete]" & _
           ",[CanAddFolder]" & _
           ",[CanEditFolder]" & _
           ",[CanDeleteFolder]" & _
           ",[CanManageTemplates]" & _
           ",[CanImport]" & _
           ",[CanExport]" & _
           ",[IsSystemAdministrator] FROM tbl_UserGroup WHERE [ID]=" & lGroupUserPK, rsGroupUser, , Not bAddState, "tbl_UserGroup"
           
    If bAddState = True Then
        Set LastRecordsetA = GetRecords("SELECT * FROM tbl_UserGroup WHERE [ID]=" & LAST_GENERATED_IDENTITY)
    Else
        If AppCurrentUser.UserGroupId = lGroupUserPK Then
            MsgBox "Changes will take effect after you restart the application.", vbInformation
        End If
        Set LastRecordsetA = rsGroupUser
    End If
    HideWaiting Me, WaitingDisplay
    Unload Me
    
    
End Sub

Private Sub Form_Load()
    
    If lGroupUserPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Set rsGroupUser = GetRecords("SELECT * FROM tbl_UserGroup WHERE [ID]=" & lGroupUserPK)
    If rsGroupUser Is Nothing Then
        MsgBox "Unable to read user group data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        BindControlToRS Me, rsGroupUser
    End If
    
    If bAddState = True Then
        Me.Caption = "Add New User Group"
        rsGroupUser.AddNew
        
        rsGroupUser.Fields("CanViewConfidential") = 0
        rsGroupUser.Fields("CanAdd") = 0
        rsGroupUser.Fields("CanEdit") = 0
        rsGroupUser.Fields("CanCheckOut") = 0
        rsGroupUser.Fields("CanDelete") = 0
        rsGroupUser.Fields("CanImport") = 0
        rsGroupUser.Fields("CanExport") = 0
        rsGroupUser.Fields("IsSystemAdministrator") = 0
        rsGroupUser.Fields("CanAddFolder") = 0
        rsGroupUser.Fields("CanEditFolder") = 0
        rsGroupUser.Fields("CanDeleteFolder") = 0
        rsGroupUser.Fields("CanManageTemplates") = 0
        
    Else
        Me.Caption = "Edit User Group"
    End If
    
    lblTitle.Caption = Me.Caption
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUserGroupAddEdit = Nothing
End Sub
