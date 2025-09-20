VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecycleBinFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Recovery"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecycleBinFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Recover Selected"
      Default         =   -1  'True
      Height          =   390
      Left            =   7950
      TabIndex        =   1
      Top             =   6000
      Width           =   1530
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   6600
      TabIndex        =   2
      Top             =   6000
      Width           =   1230
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4950
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   5
      Top             =   6825
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton cmdPDelAll 
      Caption         =   "&Permanently Delete All Folders"
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   6000
      Width           =   2430
   End
   Begin VB.CommandButton cmdPDel 
      Caption         =   " Delete Selected"
      Height          =   390
      Left            =   2700
      TabIndex        =   3
      Top             =   6000
      Width           =   1530
   End
   Begin MSComctlLib.ListView lstvExplorer 
      Height          =   4590
      Left            =   150
      TabIndex        =   0
      Top             =   1125
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   8096
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgDocIco16"
      SmallIcons      =   "imgDocIco16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Folder Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Directory"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Deleted By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Deletion Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File Contains"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDocIco16 
      Left            =   5250
      Top             =   6030
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
            Picture         =   "frmRecycleBinFolders.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -150
      TabIndex        =   7
      Top             =   5850
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1800
      TabIndex        =   8
      Top             =   7275
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Label Label2 
      Caption         =   "List of deleted folders to recover:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   825
      Width           =   3765
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Recovery"
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
      Width           =   5865
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmRecycleBinFolders.frx":0724
      Top             =   150
      Width           =   360
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
Attribute VB_Name = "frmRecycleBinFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPDel_Click()
    If lstvExplorer.ListItems.Count = 0 Then Beep: Exit Sub
    If MsgBox("Are you sure you want to permanently delete the folder named '" & lstvExplorer.SelectedItem.Text & "' and all files on it?", vbCritical + vbYesNo, "Confirm Folder Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_Folders WHERE [ID]=" & lstvExplorer.SelectedItem.Tag & "") = True Then
            HideWaiting Me, WaitingDisplay
            lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
        Else
            HideWaiting Me, WaitingDisplay
        End If
        LastGenericText = "yes"
    End If
End Sub

Private Sub cmdPDelAll_Click()

    If MsgBox("Are you sure you want to permanently delete all deleted folders and all files on it?", vbCritical + vbYesNo, "Confirm Folder Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_Folders WHERE [IsDeleted]=1") = True Then
            HideWaiting Me, WaitingDisplay
            lstvExplorer.ListItems.Clear
        Else
            HideWaiting Me, WaitingDisplay
        End If
        LastGenericText = "yes"
    End If
End Sub

Private Sub cmdSave_Click()
    If lstvExplorer.ListItems.Count = 0 Then Beep: Exit Sub
    
    RecoverSelectedFolder Val(lstvExplorer.SelectedItem.Tag), lstvExplorer.SelectedItem.Text
    
End Sub

Private Sub RecoverSelectedFolder(ByVal lFolderId As Long, ByVal folderName As String)

    ShowWaiting Me, WaitingDisplay

    Dim rsRecord  As recordset
    Set rsRecord = GetRecords("SELECT [IsDeleted],[DeletedBy],[DeletionDate] FROM tbl_Folders WHERE [ID]=" & lFolderId)
    
    rsRecord.Fields("IsDeleted") = 0
    rsRecord.Fields("DeletedBy") = ""
    rsRecord.Fields("DeletionDate") = ""
    
    If SaveRecord("", rsRecord, , True) = 1 Then
        HideWaiting Me, WaitingDisplay
        MsgBox "The folder named '" & folderName & "' has been successfully recovered.", vbInformation
        
        lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
    Else
        HideWaiting Me, WaitingDisplay
    End If
    Set rsRecord = Nothing
    
'    If ExecSQL("UPDATE tbl_Folders SET [IsDeleted]=0,[DeletedBy]='',[DeletionDate]='' WHERE [ID]=" & lFolderId & "") = True Then
'        HideWaiting Me, WaitingDisplay
'        MsgBox "The folder named '" & folderName & "' has been successfully recovered.", vbInformation
'
'        lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
'    Else
'        HideWaiting Me, WaitingDisplay
'    End If

    LastGenericText = "yes"
End Sub


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    'Load the folders
    LoadFolders
    Screen.MousePointer = vbDefault
End Sub


Public Sub LoadFolders()
    DoEvents
    
    lstvExplorer.ListItems.Clear
    
    ShowWaiting Me, WaitingDisplay
    
    Dim rsFolders As recordset
    Set rsFolders = GetRecords("SELECT [FolderName],[DirectoryName],[DeletedBy],[DeletionDate],[TotalFile],[ID] FROM vw_FoldersDeleted ORDER BY [DeletionDate] DESC")

    If Not rsFolders Is Nothing Then
        If rsFolders.RecordCount > 0 Then
            FillListView lstvExplorer, rsFolders, 5, 1, False, True, "ID"
        Else
        End If

        Set rsFolders = Nothing
    End If
    
    HideWaiting Me, WaitingDisplay
    
    
End Sub

