VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUserAddEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pnlGeneral 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3240
      Left            =   150
      ScaleHeight     =   3240
      ScaleWidth      =   4515
      TabIndex        =   13
      Top             =   1200
      Width           =   4515
      Begin VB.TextBox txtCPwd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1425
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   900
         Width           =   1965
      End
      Begin VB.TextBox txtUserName 
         DataField       =   "UserNameDec"
         Height          =   315
         Left            =   1425
         TabIndex        =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.TextBox txtNPwd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1425
         PasswordChar    =   "•"
         TabIndex        =   1
         Top             =   525
         Width           =   1965
      End
      Begin VB.TextBox txtCompleteName 
         DataField       =   "CompleteName"
         Height          =   315
         Left            =   1425
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1425
         Width           =   3015
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         Height          =   315
         Left            =   1425
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtContactNo 
         DataField       =   "ContactNo"
         Height          =   315
         Left            =   1425
         TabIndex        =   5
         Top             =   2175
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo dcGroup 
         DataField       =   "UserGroupID"
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   2520
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
         Height          =   315
         Left            =   0
         TabIndex        =   18
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Complete Name:"
         Height          =   315
         Left            =   0
         TabIndex        =   17
         Top             =   1425
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Group:"
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   2550
         Width           =   1365
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No:"
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   2175
         Width           =   1365
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3765
      Left            =   75
      TabIndex        =   12
      Top             =   750
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   6641
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "general"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2355
      TabIndex        =   7
      Top             =   4605
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3600
      TabIndex        =   8
      Top             =   4605
      Width           =   1140
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1050
      TabIndex        =   11
      Top             =   5400
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmUserAddEdit.frx":038A
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
      TabIndex        =   9
      Top             =   150
      Width           =   3690
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
Attribute VB_Name = "frmUserAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lUserPK As Long

Dim bAddState As Boolean
Dim rsUser As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(txtUserName) Then Exit Sub
    If IsControlEmpty(txtCompleteName) Then Exit Sub
    If IsControlEmpty(dcGroup) Then Exit Sub
    
    
    ShowWaiting Me, WaitingDisplay
    
    If bAddState = True Then
        If IsControlEmpty(txtNPwd) Then Exit Sub
        If IsControlEmpty(txtCPwd) Then Exit Sub
        
        'Confirm password
        If Encode(txtNPwd.Text) <> Encode(txtCPwd.Text) Then
            MsgBox "Both password must be equal.Please confirm it and try again!", vbCritical
            txtNPwd.SetFocus
            Exit Sub
        End If
        rsUser.Fields("UserPassword") = Encode(txtNPwd.Text)
    Else
        'Confirm password
        If txtNPwd.Text <> "" Or txtCPwd.Text <> "" Then
            If Encode(txtNPwd.Text) <> Encode(txtCPwd.Text) Then
                MsgBox "Both password must be equal.Please confirm it and try again!", vbCritical
                txtNPwd.SetFocus
                Exit Sub
            End If
            rsUser.Fields("UserPassword") = Encode(txtNPwd.Text)
        End If
    End If
    rsUser.Fields("UserName") = Encode(txtUserName.Text)
    rsUser.Fields("UserGroupID") = dcGroup.BoundText
    
    SyncronizeRecordsetBinding rsUser, Me
    SaveRecord "SELECT [UserNameDec],[UserName],[UserPassword],[CompleteName],[Address],[ContactNo],[UserGroupID] FROM tbl_Users WHERE [ID]=" & lUserPK, rsUser, , Not bAddState, "tbl_Users"
    
    HideWaiting Me, WaitingDisplay
    If bAddState = True Then
        Set LastRecordsetA = GetRecords("SELECT * FROM vw_Users WHERE [ID]=" & LAST_GENERATED_IDENTITY)
    Else
        Set LastRecordsetA = GetRecords("SELECT * FROM vw_Users WHERE [ID]=" & lUserPK)
    End If
    Unload Me
End Sub

Private Sub DBCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Load()
    pnlGeneral.BackColor = &H8000000F
    
    If lUserPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Set rsUser = GetRecords("SELECT * FROM tbl_Users WHERE [ID]=" & lUserPK)
    If rsUser Is Nothing Then
        MsgBox "Unable to read user data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        dcGroup.BoundColumn = "ID"
        dcGroup.ListField = "GroupName"
        Set dcGroup.RowSource = GetRecords("SELECT * FROM tbl_UserGroup ORDER BY GroupName ASC")
        
        BindControlToRS Me, rsUser
    End If
    
    If bAddState = True Then
        Me.Caption = "Add New User"
        rsUser.AddNew
    Else
        Me.Caption = "Edit User"
        If lUserPK = AppCurrentUser.UserId Then dcGroup.Enabled = False
    End If
    
    
    
    lblTitle.Caption = Me.Caption
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUserAddEdit = Nothing
End Sub
