VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Status"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManageStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   390
      Left            =   3900
      TabIndex        =   4
      Top             =   5505
      Width           =   1140
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   390
      Left            =   1350
      TabIndex        =   2
      Top             =   5505
      Width           =   1140
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   5505
      Width           =   1140
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   390
      Left            =   2625
      TabIndex        =   3
      Top             =   5505
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   5175
      TabIndex        =   5
      Top             =   5505
      Width           =   1140
   End
   Begin MSComctlLib.ImageList imgDocIco16 
      Left            =   1395
      Top             =   5220
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
            Picture         =   "frmManageStatus.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstvRecordList 
      Height          =   4470
      Left            =   90
      TabIndex        =   0
      Top             =   765
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   7885
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgDocIco32"
      SmallIcons      =   "imgDocIco16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "StatusName"
         Text            =   "Status Name"
         Object.Width           =   10231
      EndProperty
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -675
      TabIndex        =   8
      Top             =   5400
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   525
      TabIndex        =   9
      Top             =   7500
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmManageStatus.frx":0724
      Top             =   150
      Width           =   360
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Status"
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
      TabIndex        =   6
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
Attribute VB_Name = "frmManageStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Are you sure you want to delete the record named '" & lstvRecordList.SelectedItem.Text & "'?", vbCritical + vbYesNo, "Confirm Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_FileStatus WHERE [ID]=" & lstvRecordList.SelectedItem.Tag & "") = True Then
            lstvRecordList.ListItems.Remove lstvRecordList.SelectedItem.Index
            lstvRecordList.SelectedItem.Selected = True
        End If
        HideWaiting Me, WaitingDisplay
    End If
    
End Sub

Private Sub cmdEdit_Click()
    Set LastRecordsetA = Nothing
    frmManageStatusAddEdit.lPK = Val(lstvRecordList.SelectedItem.Tag)
    frmManageStatusAddEdit.Show vbModal
    
    If Not LastRecordsetA Is Nothing Then UpdateListView lstvRecordList, LastRecordsetA, 1, "ID", "StatusName", , , , , , False
    Set LastRecordsetA = Nothing
End Sub

Private Sub cmdNew_Click()
    Set LastRecordsetA = Nothing
    frmManageStatusAddEdit.lPK = 0
    frmManageStatusAddEdit.Show vbModal
    
    If Not LastRecordsetA Is Nothing Then UpdateListView lstvRecordList, LastRecordsetA, 1, "ID", "StatusName", , , , , , True
    Set LastRecordsetA = Nothing
End Sub

Private Sub cmdRefresh_Click()
    LoadGroupList
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    LoadGroupList
    Screen.MousePointer = vbDefault
End Sub


Private Sub LoadGroupList()

    ShowWaiting Me, WaitingDisplay
    
    lstvRecordList.ListItems.Clear
    
    Dim rsGroup As recordset
    Set rsGroup = GetRecords("SELECT StatusName,ID FROM tbl_FileStatus")
    
    If Not rsGroup Is Nothing Then
        If rsGroup.RecordCount > 0 Then
            DisableEditing True
            
            FillListView lstvRecordList, rsGroup, 1, 1, False, True, "ID", "StatusName"
        Else
            DisableEditing
        End If
    End If
    
    Set rsGroup = Nothing
    
    HideWaiting Me, WaitingDisplay
    
End Sub



Private Sub DisableEditing(Optional EnableEditing As Boolean)
    cmdEdit.Enabled = EnableEditing
    cmdDelete.Enabled = EnableEditing
End Sub

Private Sub lstvRecordList_DblClick()
    cmdEdit_Click
End Sub
