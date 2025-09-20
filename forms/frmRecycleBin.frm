VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecycleBin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Recovery"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecycleBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPDel 
      Caption         =   " Delete Selected"
      Height          =   390
      Left            =   2550
      TabIndex        =   4
      Top             =   6000
      Width           =   1530
   End
   Begin VB.CommandButton cmdPDelAll 
      Caption         =   "&Permanently Delete All Files"
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   6000
      Width           =   2280
   End
   Begin MSComctlLib.ListView lstvExplorer 
      Height          =   4590
      Left            =   3450
      TabIndex        =   0
      Top             =   1125
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   8096
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Original Author"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Doc. No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Doc. Index"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Doc. Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Doc. Type"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Physical Location"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Expiry"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Expiry Note"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Date Created"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Last Modified"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Created By"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Last Modified By"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Is Confidential"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Deleted By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Deletion Date"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   4950
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   7
      Top             =   6825
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   7275
      TabIndex        =   3
      Top             =   6000
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Recover Selected"
      Default         =   -1  'True
      Height          =   390
      Left            =   8625
      TabIndex        =   2
      Top             =   6000
      Width           =   1530
   End
   Begin MSComctlLib.TreeView trvFolderList 
      Height          =   4575
      Left            =   150
      TabIndex        =   1
      Top             =   1125
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   619
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgTVFolderList"
      Appearance      =   1
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
            Picture         =   "frmRecycleBin.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDocIco24 
      Left            =   5925
      Top             =   6030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecycleBin.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDocIco32 
      Left            =   6555
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecycleBin.frx":0E9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTVFolderList 
      Left            =   4575
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecycleBin.frx":1B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecycleBin.frx":1F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecycleBin.frx":22AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -150
      TabIndex        =   10
      Top             =   5850
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1800
      TabIndex        =   11
      Top             =   7275
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmRecycleBin.frx":26EC
      Top             =   150
      Width           =   360
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "File Recovery"
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
      TabIndex        =   8
      Top             =   150
      Width           =   5865
   End
   Begin VB.Label Label2 
      Caption         =   "List of deleted files to recover:"
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
      TabIndex        =   6
      Top             =   825
      Width           =   3765
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
Attribute VB_Name = "frmRecycleBin"
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
    If MsgBox("Are you sure you want to permanently delete the file named '" & lstvExplorer.SelectedItem.Text & "'?", vbCritical + vbYesNo, "Confirm File Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_Files WHERE [ID]=" & lstvExplorer.SelectedItem.Tag & "") = True Then
            HideWaiting Me, WaitingDisplay
            lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
        Else
            HideWaiting Me, WaitingDisplay
        End If
        LastGenericText = "yes"
    End If
End Sub

Private Sub cmdPDelAll_Click()

    If MsgBox("Are you sure you want to permanently delete all deleted files?", vbCritical + vbYesNo, "Confirm File Deletion") = vbYes Then
        ShowWaiting Me, WaitingDisplay
        If ExecSQL("DELETE tbl_Files WHERE [IsDeleted]=1") = True Then
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
    
    RecoverSelectedFile Val(lstvExplorer.SelectedItem.Tag), lstvExplorer.SelectedItem.Text
    
End Sub

Private Sub RecoverSelectedFile(ByVal lFileId As Long, ByVal fileName As String)

    ShowWaiting Me, WaitingDisplay
    
    Dim rsRecord As recordset
    Set rsRecord = GetRecords("SELECT [IsDeleted],[DeletedBy],[DeletionDate] FROM tbl_Files WHERE [ID]=" & lFileId)
    
    rsRecord.Fields("IsDeleted") = 0
    rsRecord.Fields("DeletedBy") = ""
    rsRecord.Fields("DeletionDate") = ""
    
    If SaveRecord("", rsRecord, , True) = 1 Then
        HideWaiting Me, WaitingDisplay
        MsgBox "The file named '" & fileName & "' has been successfully recovered.", vbInformation
        
        lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
    Else
        HideWaiting Me, WaitingDisplay
    End If
    Set rsRecord = Nothing

'    If ExecSQL("UPDATE tbl_Files SET [IsDeleted]=0,[DeletedBy]='',[DeletionDate]='' WHERE [ID]=" & lFileId & "") = True Then
'        HideWaiting Me, WaitingDisplay
'        MsgBox "The file named '" & fileName & "' has been successfully recovered.", vbInformation
'
'        lstvExplorer.ListItems.Remove (lstvExplorer.SelectedItem.Index)
'    Else
'        HideWaiting Me, WaitingDisplay
'    End If

    LastGenericText = "yes"
End Sub


Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    LastUseFileId = ""
    LastUseFileNamePath = ""
    
    'Load the folders
    LoadDirectory ""
    trvFolderList.SelectedItem = trvFolderList.Nodes(2)
    LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
    trvFolderList.SelectedItem = trvFolderList.Nodes(1)
    LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
    
    
    LoadFiles Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
    Screen.MousePointer = vbDefault
End Sub


Private Sub LoadDirectory(ByVal DirectoryName As String, Optional ParentKey As String, Optional ForceToExpand As Boolean, Optional IsNewlyInserted As Boolean)
    DoEvents
    
    On Error GoTo err
    
    ShowWaiting Me, WaitingDisplay
    
    
    Dim rsFolder As recordset
    
    If ParentKey = "" Then
        Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
    Else
        If IsNewlyInserted = True Then
            Set rsFolder = GetRecords("SELECT TOP 1 * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [ID] DESC,[DirectoryName],[FolderName] ASC")
        Else
            Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
        End If
    End If
    
'    If ParentKey = "" Then
'        Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE FolderName='Templates' ORDER BY [DirectoryName],[FolderName] ASC")
'    Else
'        If IsNewlyInserted = True Then
'            Set rsFolder = GetRecords("SELECT TOP 1 * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [ID] DESC,[DirectoryName],[FolderName] ASC")
'        Else
'            Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
'        End If
'    End If
    
    'Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
    
    If Not rsFolder Is Nothing Then
        If rsFolder.RecordCount > 0 Then
            Dim tmpNode As Node
            
            rsFolder.MoveFirst
            
            Do While Not rsFolder.EOF
                If DirectoryName = "" Then
                    Set tmpNode = trvFolderList.Nodes.Add(, , "ID:" & rsFolder.Fields("ID"), rsFolder.Fields("FolderName"), 3, 3)
                    tmpNode.Bold = True
                    tmpNode.Expanded = True
                    tmpNode.Selected = True
                    
                Else
                    Set tmpNode = trvFolderList.Nodes.Add(ParentKey, tvwChild, "ID:" & rsFolder.Fields("ID"), rsFolder.Fields("FolderName"), 2, 1)
                    tmpNode.Expanded = False
                    If ForceToExpand = True Then
                        tmpNode.Selected = True
                        tmpNode.Parent.Expanded = ForceToExpand
                    End If
                    
                End If
                
                tmpNode.Tag = "ID:" & rsFolder.Fields("ID")
                
                If DirectoryName <> "" Then
                    'LoadDirectory rsFolder.Fields("DirectoryName") & rsFolder.Fields("FolderName") & "\", "ID:" & rsFolder.Fields("ID")
                    If rsFolder.Fields("NoOfSubFolders") > 0 Then
                        Set tmpNode = trvFolderList.Nodes.Add(tmpNode.Tag, tvwChild, , "Loading...", 0, 0)
                        tmpNode.Tag = "loading"
                    End If
                End If
                
                rsFolder.MoveNext
            Loop
            
            Set tmpNode = Nothing
        End If
        
    End If
    
    Set rsFolder = Nothing
    
    HideWaiting Me, WaitingDisplay
    Exit Sub
err:
    If err.Number = 35602 Or err.Number = 91 Or err.Number = 35605 Then
        Resume Next
    Else
        HideWaiting Me, WaitingDisplay
        'InputBox err.Description, "", err.Number
    End If
End Sub


Public Sub LoadFiles(ByVal folderId As Long)
    DoEvents
    
    lstvExplorer.ListItems.Clear
    
    ShowWaiting Me, WaitingDisplay
    
    Dim rsDocFiles As recordset
    Set rsDocFiles = GetRecords("SELECT [FileName],[Title],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[FileSize],[AlertDate],[AlertNote],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy],[IsConfidential],[DeletedBy],[DeletionDate],[ID],[Description] FROM vw_FileInfoOnlyDeleted WHERE [FolderID]=" & folderId & "")
    

    If Not rsDocFiles Is Nothing Then
        If rsDocFiles.RecordCount > 0 Then
            FillListView lstvExplorer, rsDocFiles, 18, -1, False, True, "ID", "Description", "IsConfidential", "FileName", imgDocIco32, imgDocIco16, picTemp
        Else
        End If

        Set rsDocFiles = Nothing
    End If
    
    HideWaiting Me, WaitingDisplay
    
    
End Sub

Private Sub trvFolderList_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Node.EnsureVisible
    
    LoadFiles Replace(Node.Tag, "ID:", "")
End Sub

Private Sub trvFolderList_NodeClick(ByVal Node As MSComctlLib.Node)
  
    LoadFiles Replace(Node.Tag, "ID:", "")
End Sub

Private Sub trvFolderList_Expand(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Node.EnsureVisible
    LoadFiles Replace(Node.Tag, "ID:", "")

    If Node.children > 0 Then
        If Node.Child.Tag = "loading" Then
            'Remove the temporary child
            trvFolderList.Nodes.Remove (Node.Child.Index)

            LoadDirectory Node.FullPath & "\", Node.Tag
        End If
    End If
End Sub

