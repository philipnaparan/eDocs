VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DD53C357-B171-4403-B656-34DFAB17A8B1}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmFileManager 
   Caption         =   "e-Docs vr 2.0 "
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   2325
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Timer tmrActiveControlListener 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3645
      Top             =   2940
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   5130
      Top             =   2790
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   30
      Bmp:1           =   "frmFileManager.frx":2CFA
      Key:1           =   "#mnuAbout"
      Bmp:2           =   "frmFileManager.frx":3122
      Key:2           =   "#mnuHelpContent"
      Bmp:3           =   "frmFileManager.frx":354A
      Key:3           =   "#mnuManUsr"
      Bmp:4           =   "frmFileManager.frx":3972
      Key:4           =   "#mnuExit"
      Bmp:5           =   "frmFileManager.frx":3D9A
      Key:5           =   "#mnuNew"
      Bmp:6           =   "frmFileManager.frx":41C2
      Key:6           =   "#mnuEdit"
      Bmp:7           =   "frmFileManager.frx":45EA
      Mask:7          =   16711935
      Key:7           =   "#mnuDel"
      Bmp:8           =   "frmFileManager.frx":493C
      Mask:8          =   16711935
      Key:8           =   "#mnuRefresh"
      Bmp:9           =   "frmFileManager.frx":4C8E
      Mask:9          =   16711935
      Key:9           =   "#mnuView"
      Bmp:10          =   "frmFileManager.frx":4FE0
      Mask:10         =   16711935
      Key:10          =   "#mnuImportDoc"
      Bmp:11          =   "frmFileManager.frx":5332
      Key:11          =   "#mnuOpenDoc"
      Bmp:12          =   "frmFileManager.frx":575A
      Key:12          =   "#mnuBatchUpdate"
      Bmp:13          =   "frmFileManager.frx":5B82
      Mask:13         =   16711935
      Key:13          =   "#mnuExportDoc"
      Bmp:14          =   "frmFileManager.frx":5ED4
      Key:14          =   "#mnuNewFolder"
      Bmp:15          =   "frmFileManager.frx":62FC
      Key:15          =   "#mnuRefreshFolderList"
      Bmp:16          =   "frmFileManager.frx":6724
      Key:16          =   "#mnuRenFolder"
      Bmp:17          =   "frmFileManager.frx":6B4C
      Key:17          =   "#mnuDelFolder"
      Bmp:18          =   "frmFileManager.frx":6F74
      Mask:18         =   16711935
      Key:18          =   "#mnuReminder"
      Bmp:19          =   "frmFileManager.frx":72C6
      Key:19          =   "#mnuDuplicate"
      Bmp:20          =   "frmFileManager.frx":76EE
      Key:20          =   "#mnuFolderInfo"
      Bmp:21          =   "frmFileManager.frx":7B16
      Key:21          =   "#mnuManageUserGroup"
      Bmp:22          =   "frmFileManager.frx":7F3E
      Key:22          =   "#mnuRecycleBin"
      Bmp:23          =   "frmFileManager.frx":8366
      Key:23          =   "#mnuMetaData"
      Bmp:24          =   "frmFileManager.frx":878E
      Key:24          =   "#mnuDocStatusSet"
      Bmp:25          =   "frmFileManager.frx":8BB6
      Key:25          =   "#mnuDocTypeSet"
      Bmp:26          =   "frmFileManager.frx":8FDE
      Key:26          =   "#mnuCheckIn"
      Bmp:27          =   "frmFileManager.frx":9406
      Key:27          =   "#mnuCheckOut"
      Bmp:28          =   "frmFileManager.frx":982E
      Key:28          =   "#mnuFolderRecover"
      Bmp:29          =   "frmFileManager.frx":9C56
      Key:29          =   "#mnuFileRecover"
      Bmp:30          =   "frmFileManager.frx":A07E
      Key:30          =   "#mnuDocxStat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTVFolderList 
      Left            =   4200
      Top             =   5355
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
            Picture         =   "frmFileManager.frx":A4A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":A840
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":ABDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTBNorm 
      Left            =   4770
      Top             =   3450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":B01A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":B794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":BF0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":C688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":CE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":D516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":DC10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":E30A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":EA04
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":F0FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":F812
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":FF26
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1063A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":10D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":1142E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":11BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileManager.frx":12322
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbShortcut 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTBNorm"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newfolder"
            Object.ToolTipText     =   "New Folder"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "renfolder"
            Object.ToolTipText     =   "Rename Folder"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delfolder"
            Object.ToolTipText     =   "Delete Folder"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refreshfolder"
            Object.ToolTipText     =   "Resfresh Folder List"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New File"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Edit File"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "duplicate"
            Object.ToolTipText     =   "Duplicate"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "import"
            Object.ToolTipText     =   "Import File(s)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "export"
            Object.ToolTipText     =   "Export Selected File(s)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "checkin"
            Object.ToolTipText     =   "Check-In"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "checkout"
            Object.ToolTipText     =   "Check-Out"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh File List"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "view"
            Object.ToolTipText     =   "Views"
            ImageIndex      =   16
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Icon"
                  Text            =   "Icon"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SmallIcon"
                  Text            =   "Small Icon"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List"
                  Text            =   "List"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Details"
                  Text            =   "Details"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   7740
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8890
            Picture         =   "frmFileManager.frx":12A1C
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5821
            MinWidth        =   5821
            Picture         =   "frmFileManager.frx":12DB6
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8100
      Top             =   2025
   End
   Begin VB.PictureBox picLeftContainer 
      Align           =   3  'Align Left
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   0
      ScaleHeight     =   7200
      ScaleWidth      =   3285
      TabIndex        =   1
      Top             =   480
      Width           =   3345
      Begin MSComctlLib.TreeView trvFolderList 
         Height          =   6225
         Left            =   45
         TabIndex        =   2
         Top             =   315
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   10980
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   619
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgTVFolderList"
         Appearance      =   0
         OLEDragMode     =   1
         OLEDropMode     =   1
      End
      Begin VB.Label lblFolderName 
         BackColor       =   &H8000000D&
         Caption         =   " Folders"
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
         TabIndex        =   6
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.PictureBox picSeparator 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   3345
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7260
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   480
      Width           =   50
   End
   Begin eDocs.ctrlDocViewer docViewer 
      Height          =   8160
      Left            =   3600
      TabIndex        =   3
      Top             =   750
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   14393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogOutUser 
         Caption         =   "Log-out User"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuFolderAction 
      Caption         =   "&Folders"
      Begin VB.Menu mnuNewFolder 
         Caption         =   "&New Folder"
      End
      Begin VB.Menu mnuRenFolder 
         Caption         =   "&Edit Folder"
      End
      Begin VB.Menu mnuDelFolder 
         Caption         =   "&Delete Folder"
      End
      Begin VB.Menu mnuRefreshFolderList 
         Caption         =   "&Refresh Folder List"
      End
      Begin VB.Menu mnuFolderActionSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFolderInfo 
         Caption         =   "Folder &Info."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "F&iles"
      Begin VB.Menu mnuOpenDoc 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenWith 
         Caption         =   "Open &with..."
      End
      Begin VB.Menu mnuActionSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New File"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit File"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuBatchUpdate 
         Caption         =   "&Batch Update"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDuplicate 
         Caption         =   "D&uplicate"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete File"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuImportDoc 
         Caption         =   "&Import File(s)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExportDoc 
         Caption         =   "&Export Selected File(s)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCheckIn 
         Caption         =   "&Check In"
      End
      Begin VB.Menu mnuCheckOut 
         Caption         =   "C&heck Out"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuMarkSelAs 
         Caption         =   "&Mark Selected As..."
         Begin VB.Menu mnuMarkConfi 
            Caption         =   "Mark As &Confidential"
         End
         Begin VB.Menu mnuMarkNonConfi 
            Caption         =   "Mark As &Non-Confidential"
         End
      End
      Begin VB.Menu mnuReminder 
         Caption         =   "&View Reminders"
         Begin VB.Menu mnuViewExpiredDocs 
            Caption         =   "&View Expired Alerts"
         End
         Begin VB.Menu mnuActionSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewNext7 
            Caption         =   "View Alerts In Next 7 Days"
         End
         Begin VB.Menu mnuViewNext15 
            Caption         =   "View Alerts In Next 15 Days"
         End
         Begin VB.Menu mnuViewNext30 
            Caption         =   "View Alerts In Next 30 Days"
         End
         Begin VB.Menu mnuViewNext60 
            Caption         =   "View Alerts In Next 60 Days"
         End
         Begin VB.Menu mnuViewNext90 
            Caption         =   "View Alerts In Next 90 Days"
         End
      End
      Begin VB.Menu mnuActionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh File List"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Files By"
         Begin VB.Menu mnuIcon 
            Caption         =   "Icons"
         End
         Begin VB.Menu mnuSmallIco 
            Caption         =   "Small Icon"
         End
         Begin VB.Menu mnuList 
            Caption         =   "List"
         End
         Begin VB.Menu mnuDetails 
            Caption         =   "Details"
         End
      End
   End
   Begin VB.Menu mnuAdminTools 
      Caption         =   "&Administrative Tools"
      Begin VB.Menu mnuManUsr 
         Caption         =   "&Manage Users"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuManageUserGroup 
         Caption         =   "Manage User &Group"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSetupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFolderRecover 
         Caption         =   "Folder Recovery"
      End
      Begin VB.Menu mnuFileRecover 
         Caption         =   "File Recovery"
      End
      Begin VB.Menu FileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMetaData 
         Caption         =   "E&xtra Property Setup"
      End
      Begin VB.Menu mnuDocStatusSet 
         Caption         =   "&Status Setup"
      End
      Begin VB.Menu mnuDocTypeSet 
         Caption         =   "Document &Type Setup"
      End
      Begin VB.Menu FileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDocxStat 
         Caption         =   "Docx System Status"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "&Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuShowSplash 
         Caption         =   "&Show Splash Screen"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About e-Docs..."
      End
   End
End
Attribute VB_Name = "frmFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'For point api function
Dim cursor_pos As POINTAPI

Dim resize_down     As Boolean
Dim pos_num         As Integer

Dim strActiveControl As String
Dim iFileviewcode As Integer

Dim strPrevActiveCtrl As String
Dim iPrevDocSelTabIndex As Integer


Public Sub InitializeForm()
    If LoginOk = False Then
        Exit Sub
        Unload Me
    Else
        tmrActiveControlListener.Enabled = True
    End If
    
    mnuAdminTools.Visible = AppCurrentUser.bIsSysAdmin
    
    'Load the folders
    LoadDirectory ""
    If AppCurrentUser.bCanManageTemplates = True Then
        trvFolderList.SelectedItem = trvFolderList.Nodes(2)
        LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
    Else
        trvFolderList.Nodes.Remove (trvFolderList.Nodes(2).Index)
    End If
    trvFolderList.SelectedItem = trvFolderList.Nodes(1)
    LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
  
    
    'Display the files
    CheckFolderLeverRestrictions Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
    docViewer.folderId = Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
        
    
    'Display the current user
    statBar.Panels(3).Text = "Current User: " & AppCurrentUser.CompleteName
    statBar.Panels(3).ToolTipText = statBar.Panels(3).Text
    
    frmReminders.Show vbModal
    Select Case LastViewedReminderOption
        Case "exprd": mnuViewExpiredDocs_Click
        Case "next7": mnuViewNext7_Click
        Case "next15": mnuViewNext15_Click
        Case "next30": mnuViewNext30_Click
        Case "next60": mnuViewNext60_Click
        Case "next90": mnuViewNext90_Click
    End Select
    
    docViewer.SetFocus
End Sub

Private Sub docViewer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuAction
    End If
End Sub

Private Sub docViewer_NavChange(NaveInfo As String)
    statBar.Panels(2).Text = NaveInfo
End Sub

Private Sub Form_Load()
    
    iFileviewcode = GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""fileviewcode""]", "value", App.Path & "\settings\AppSettings.xml")
    Me.WindowState = GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""winstate""]", "value", App.Path & "\settings\AppSettings.xml")
    ChangeExplorerView
    
    If IsDemo = True Then
        Me.Caption = "e-Docs vr 2.0  [TRIAL]"
    Else
        Me.Caption = "e-Docs vr 2.0  [REGISTERED]"
    End If
    
End Sub

Private Sub ChangeExplorerView()
    docViewer.ChangeView iFileviewcode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    tmrActiveControlListener.Enabled = False
    If LoginOk = True Then
        If MsgBox("This will close the application.Do you want to proceed?", vbExclamation + vbYesNo) = vbNo Then
            Cancel = 1
            tmrActiveControlListener.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    PositionDocViewer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""fileviewcode""]", "value", iFileviewcode, App.Path & "\settings\AppSettings.xml"
    UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""winstate""]", "value", Me.WindowState, App.Path & "\settings\AppSettings.xml"
    
    Shell App.Path & "\temp\TempEraser.exe " & App.Path & "\temp\"
    
    Set frmFileManager = Nothing
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBatchUpdate_Click()
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    
    docViewer.ProccessCommand "UPDATE_BATCH"
End Sub

Private Sub mnuCheckIn_Click()
    docViewer.CheckIn
End Sub

Private Sub mnuCheckOut_Click()
    docViewer.CheckOut
End Sub

Private Sub mnuDel_Click()
    If AppCurrentUser.bCanDelete = False Then PrompAccessDenied: Exit Sub
    
    DeleteSelectedDoc

End Sub

Private Sub DeleteSelectedFolder()
    Dim strDelMsg As String
    
    If trvFolderList.SelectedItem.children > 0 Then
        MsgBox "This folder have a sub-folder(s), please delete the sub-folder(s) first before deleting this folder.", vbCritical
        Exit Sub
    Else
        strDelMsg = "Are you sure you want to delete the folder named '" & trvFolderList.SelectedItem.Text & "'?" & vbCrLf & vbCrLf & _
              "Warning: All files on it will be deleted also."
    End If
    
    
    
    If MsgBox(strDelMsg, vbCritical + vbYesNo, "Confirm Folder Deletion") = vbYes Then
    
        Dim rsFolder As recordset
        
        Set rsFolder = GetRecords("SELECT [IsDeleted],[DeletedBy],[DeletionDate] FROM tbl_Folders WHERE [ID]=" & Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", "")))
        
        rsFolder.Fields("IsDeleted") = 1
        rsFolder.Fields("DeletedBy") = AppCurrentUser.CompleteName
        rsFolder.Fields("DeletionDate") = Now
        
        If SaveRecord("", rsFolder, , True) = 1 Then
            trvFolderList.Nodes.Remove (trvFolderList.SelectedItem.Index)
            CheckFolderLeverRestrictions Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
            docViewer.folderId = Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
        End If
        Set rsFolder = Nothing
    End If
    
End Sub

Private Sub DeleteSelectedDoc()
        
    docViewer.ProccessCommand "DELETE_DOC"
    
End Sub

Private Sub mnuDelFolder_Click()
    If AppCurrentUser.bCanDeleteFolder = False Then PrompAccessDenied: Exit Sub
    If trvFolderList.SelectedItem.Text = trvFolderList.SelectedItem.FullPath Then Beep: Exit Sub
    
    If CurrentFolderAccess.folderId = Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", "")) And CurrentFolderAccess.bDenyFolderDelete = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DeleteSelectedFolder
End Sub

Private Sub mnuDetails_Click()
    iFileviewcode = 4
    ChangeExplorerView
End Sub

Private Sub mnuDocStatusSet_Click()
    frmManageStatus.Show vbModal
End Sub

Private Sub mnuDocTypeSet_Click()
    frmManageDocumentType.Show vbModal
End Sub

Private Sub mnuDocxStat_Click()
    ShowWaiting Me, WaitingDisplay
    
    
    On Error Resume Next
    Dim totalAFile As Long
    Dim totalAFileSize As Long
    Dim totalAFolder As Long
    
    Dim totalDFile As Long
    Dim totalDFileSize As Long
    Dim totalDFolder As Long
    
    totalAFile = GetRecordCount("SELECT ID FROM tbl_Files WHERE [IsDeleted]=0")
    totalAFileSize = GetValueAt("SELECT SUM(FileSizeNo) AS 'Total' FROM tbl_Files WHERE [IsDeleted]=0", "Total")
    totalAFolder = GetRecordCount("SELECT ID FROM tbl_Folders WHERE [IsDeleted]=0")
    
    totalDFile = GetRecordCount("SELECT ID FROM tbl_Files WHERE [IsDeleted]=1")
    totalDFileSize = GetValueAt("SELECT SUM(FileSizeNo) AS 'Total' FROM tbl_Files WHERE [IsDeleted]=1", "Total")
    totalDFolder = GetRecordCount("SELECT ID FROM tbl_Folders WHERE [IsDeleted]=1")
        
    Dim msg As String
    
    msg = vbCrLf & _
        "[Active File and Folders]" & vbCrLf & vbCrLf & _
        "Total File(s): " & totalAFile & vbCrLf & _
        "Total Size Of File(s): " & GetFileSizeInfo(totalAFileSize) & vbCrLf & _
        "Total Folder(s): " & totalAFolder & vbCrLf & _
        vbCrLf & vbCrLf & _
        "[Recycle Bin]" & vbCrLf & vbCrLf & _
        "Total Deleted File(s): " & totalDFile & vbCrLf & _
        "Total Size Of Deleted File(s): " & GetFileSizeInfo(totalDFileSize) & vbCrLf & _
        "Total Deleted Folder(s): " & totalDFolder & vbCrLf & vbCrLf & vbCrLf
    
    HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
    
    MsgBox msg, vbInformation, "e-Docs System Status"
    
    
End Sub

Private Sub mnuDuplicate_Click()
    If AppCurrentUser.bCanAdd = False Then PrompAccessDenied: Exit Sub
    
    docViewer.ProccessCommand "DUPLICATE_DOC"
End Sub

Private Sub mnuEdit_Click()
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    
    EditSelectedDoc

End Sub

Public Sub RenameSelNode(ByVal NewName As String, ByVal folderId As String)
    
    UpdateSubDir folderId, trvFolderList.SelectedItem.Parent.FullPath & "\" & NewName & "\"
    trvFolderList.SelectedItem.Text = NewName
    
End Sub

Public Function IsDirectoryExist(ByVal folderName As String, Optional IsUpdateMode As Boolean) As Boolean
    Dim dirName As String
    If IsUpdateMode = True Then
        dirName = trvFolderList.SelectedItem.Parent.FullPath
    Else
        dirName = trvFolderList.SelectedItem.FullPath
    End If
     If GetRecordCount("SELECT [ID] FROM tbl_Folders WHERE [DirectoryName]='" & dirName & "\' AND [FolderName]='" & folderName & "'") > 0 Then
        MsgBox "Folder named '" & folderName & "' is already exist.", vbCritical
        IsDirectoryExist = True
    End If
End Function


Private Sub EditSelectedFolder()
    
    On Error GoTo err
    
    LastGenericText = ""
    
    Set LastRecordsetA = Nothing
    
    frmFolderEdit.lFolderId = Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", ""))
    frmFolderEdit.Show vbModal
    
    If LastGenericText = "HaveChanges" Then
        CheckFolderLeverRestrictions Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", ""))
    End If
       
err:

End Sub


Private Function UpdateSubDir(ByVal folderId As String, ByVal NewDirName As String) As Boolean

    On Error GoTo err
    Dim rsDir As recordset
    Dim rsTemp As String
    Dim currDir As String
    Set rsDir = GetRecords("SELECT * FROM tbl_Folders WHERE [ParentFolderID]=" & folderId & "")

    If Not rsDir Is Nothing Then
        If rsDir.RecordCount > 0 Then
            rsDir.MoveFirst
            Do While Not rsDir.EOF
                currDir = rsDir.Fields("DirectoryName") & rsDir.Fields("FolderName") & "\"
                If ExecSQL("UPDATE tbl_Folders SET [DirectoryName]='" & NewDirName & "' WHERE [ID]=" & rsDir.Fields("ID")) = True Then
                    If GetRecordCount("SELECT [ID] FROM tbl_Folders WHERE [ParentFolderID]=" & rsDir.Fields("ID")) > 0 Then
                        'InputBox "", , NewDirName & " - " & currDir
                        UpdateSubDir rsDir.Fields("ID"), currDir
                        
                    End If
                End If
                rsDir.MoveNext
            Loop
        End If
    End If
    Set rsDir = Nothing

    UpdateSubDir = True
    Exit Function
err:
    UpdateSubDir = False
End Function

Private Sub EditSelectedDoc()
    docViewer.ProccessCommand "EDIT_DOC"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExportDoc_Click()
    If AppCurrentUser.bCanExport = False Then PrompAccessDenied: Exit Sub
    
    docViewer.ProccessCommand "EXPORT_DOC"
End Sub

Private Sub mnuFileRecover_Click()
    LastGenericText = ""
    frmRecycleBin.Show vbModal
    
    If LastGenericText = "yes" Then RefreshDocList
    LastGenericText = ""
End Sub

Private Sub mnuFolderInfo_Click()
    '
End Sub

Private Sub mnuFolderRecover_Click()
    LastGenericText = ""
    frmRecycleBinFolders.Show vbModal
    
    If LastGenericText = "yes" Then
        On Error Resume Next
        trvFolderList.Nodes.Clear
        docViewer.Reset
        
        'Load the folders
        LoadDirectory ""
        If AppCurrentUser.bCanManageTemplates = True Then
            trvFolderList.SelectedItem = trvFolderList.Nodes(2)
            LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
        Else
            trvFolderList.Nodes.Remove (trvFolderList.Nodes(2).Index)
        End If
        trvFolderList.SelectedItem = trvFolderList.Nodes(1)
        LoadDirectory trvFolderList.SelectedItem.Text & "\", trvFolderList.SelectedItem.Tag
        
        
        'Display the files
        CheckFolderLeverRestrictions Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
        docViewer.folderId = Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
    End If
    LastGenericText = ""
End Sub

Private Sub mnuHelpContent_Click()
    On Error Resume Next
    OpenURL App.Path & "\Help.doc", Me.hwnd
End Sub

Private Sub mnuIcon_Click()
    iFileviewcode = 1
    ChangeExplorerView
End Sub

Private Sub mnuImportDoc_Click()
    If AppCurrentUser.bCanImport = False Then PrompAccessDenied: Exit Sub
    
    docViewer.ProccessCommand "IMPORT_DOC"
End Sub

Private Sub mnuList_Click()
    iFileviewcode = 3
    ChangeExplorerView
End Sub

Private Sub mnuLogOutUser_Click()
    LoginOk = False
    
    frmLogin.Show vbModal
    
    If LoginOk = False Then
        Unload Me
    Else
        docViewer.Reset
        trvFolderList.Nodes.Clear
        
        InitializeForm
    End If
End Sub

Private Sub mnuManageUserGroup_Click()
    frmManageUserGroup.Show vbModal
End Sub

Private Sub mnuManUsr_Click()

    'If LCase(AppCurrentUser.UserType) <> "administrator" Then PrompAccessDenied: Exit Sub
    frmManageUser.Show vbModal
    
End Sub

Private Sub mnuMarkConfi_Click()
    docViewer.ProccessCommand "MARK_CONFI"
End Sub

Private Sub mnuMarkNonConfi_Click()
    docViewer.ProccessCommand "MARK_NONCONFI"
End Sub

Private Sub mnuMetaData_Click()
    frmManageExtraProperties.Show vbModal
End Sub

Private Sub mnuNew_Click()
    If AppCurrentUser.bCanAdd = False Then PrompAccessDenied: Exit Sub

    CreateNewDoc
End Sub

Public Function IsFolderExist(ByVal folderName As String) As Boolean
    Dim retVal As Boolean

    If IsNodePathExist(trvFolderList.Nodes, trvFolderList.SelectedItem.FullPath & "\" & folderName) = True Then
        MsgBox "The folder named '" & folderName & "' is already exist.", vbCritical
        retVal = True
    End If
    
    IsFolderExist = retVal
End Function


Private Sub CreateNewFolder()
    On Error GoTo err
    
    
    Set LastRecordsetA = Nothing
    
    frmFolderAdd.lParentPK = Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", ""))
    frmFolderAdd.strDirName = trvFolderList.SelectedItem.FullPath & "\"
    frmFolderAdd.Show vbModal
    
    If Not LastRecordsetA Is Nothing Then
        'If successfully save then update the folder list
        LoadDirectory trvFolderList.SelectedItem.FullPath & "\", trvFolderList.SelectedItem.Tag, True, True
'        trvFolderList.SelectedItem.Selected = True
'        trvFolderList.SelectedItem.EnsureVisible
        
        statBar.Panels(1).Text = trvFolderList.SelectedItem.FullPath
    
        CheckFolderLeverRestrictions Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
        docViewer.folderId = Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
    End If
    
err:
End Sub

Private Sub CreateNewDoc()
    docViewer.ProccessCommand "NEW_DOC"
End Sub

Private Sub mnuNewFolder_Click()
    If AppCurrentUser.bCanAddFolder = False Then PrompAccessDenied: Exit Sub

    CreateNewFolder
End Sub

Private Sub mnuOpenDoc_Click()
    docViewer.ProccessCommand "OPEN_DOC"
End Sub

Private Sub mnuOpenWith_Click()
    docViewer.ProccessCommand "OPEN_WITH"
End Sub

Private Sub mnuRecycleBin_Click()
    
End Sub

Private Sub mnuRefresh_Click()
    RefreshDocList
End Sub

Private Sub RefreshFolderList()
    LoadDirectory trvFolderList.SelectedItem.FullPath & "\", trvFolderList.SelectedItem.Tag, True
End Sub

Private Sub RefreshDocList()
    docViewer.ProccessCommand "REFRESH_DOC"
End Sub

Private Sub mnuRefreshFolderList_Click()
    RefreshFolderList
End Sub

Private Sub mnuRenFolder_Click()
    If AppCurrentUser.bCanEditFolder = False Then PrompAccessDenied: Exit Sub
    If trvFolderList.SelectedItem.Text = trvFolderList.SelectedItem.FullPath Then Beep: Exit Sub
    
    If CurrentFolderAccess.folderId = Val(Replace(trvFolderList.SelectedItem.Tag, "ID:", "")) And CurrentFolderAccess.bDenyFolderEdit = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    EditSelectedFolder
End Sub

Private Sub mnuSelAll_Click()
    docViewer.ProccessCommand "SELECT_ALL"
End Sub

Private Sub mnuShowSplash_Click()
    frmSplash.bForDisplay = True
    frmSplash.Show vbModal
End Sub

Private Sub mnuSmallIco_Click()
    iFileviewcode = 2
    ChangeExplorerView
End Sub

Private Sub mnuViewExpiredDocs_Click()
    docViewer.ViewExpiry CDate("1/1/1900"), Date
End Sub

Private Sub mnuViewNext15_Click()
    docViewer.ViewExpiry Date, DateAdd("d", 15, Date)
End Sub

Private Sub mnuViewNext30_Click()
    docViewer.ViewExpiry Date, DateAdd("d", 30, Date)
End Sub

Private Sub mnuViewNext60_Click()
    docViewer.ViewExpiry Date, DateAdd("d", 60, Date)
End Sub

Private Sub mnuViewNext7_Click()
    docViewer.ViewExpiry Date, DateAdd("d", 7, Date)
End Sub

Private Sub mnuViewNext90_Click()
    docViewer.ViewExpiry Date, DateAdd("d", 90, Date)
End Sub

Private Sub picSeparator_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        tmrResize.Enabled = True
        resize_down = True
    End If
End Sub

Private Sub picSeparator_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tmrResize.Enabled = False
        resize_down = False
    End If
End Sub

Private Sub picLeftContainer_Resize()
    On Error Resume Next
    
    lblFolderName.Top = 0
    lblFolderName.Width = picLeftContainer.Width
    
    trvFolderList.Top = lblFolderName.Height + 50
    
    
    trvFolderList.Width = picLeftContainer.Width - (trvFolderList.Left + 100)
    trvFolderList.Height = picLeftContainer.Height - (trvFolderList.Top + 100)
End Sub

Private Sub tbShortcut_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "newfolder"
            mnuNewFolder_Click
        Case "renfolder"
            mnuRenFolder_Click
        Case "delfolder"
            mnuDelFolder_Click
        Case "refreshfolder"
            mnuRefreshFolderList_Click
        Case "folderinfo"
            mnuFolderInfo_Click
        Case "open"
            mnuOpenDoc_Click
        Case "new"
            mnuNew_Click
        Case "edit"
            mnuEdit_Click
        Case "import"
            mnuImportDoc_Click
        Case "delete"
            mnuDel_Click
        Case "refresh"
            mnuRefresh_Click
        Case "export"
            mnuExportDoc_Click
        Case "duplicate"
            mnuDuplicate_Click
        Case "checkin"
            mnuCheckIn_Click
        Case "checkout"
            mnuCheckOut_Click
      
    End Select
End Sub

Private Sub tbShortcut_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    iFileviewcode = ButtonMenu.Index
    ChangeExplorerView
End Sub

Private Sub tmrActiveControlListener_Timer()
    On Error Resume Next
    
    If (Me.ActiveControl.Name <> "docViewer" And Me.ActiveControl.Name <> "trvFolderList") And docViewer.HaveChanges = False Then
        Exit Sub
    End If
    
    If strPrevActiveCtrl <> Me.ActiveControl.Name Or iPrevDocSelTabIndex <> docViewer.GetSelectedTabIndex Or docViewer.HaveChanges = True Then
        strPrevActiveCtrl = Me.ActiveControl.Name
        iPrevDocSelTabIndex = docViewer.GetSelectedTabIndex
        docViewer.HaveChanges = False
    Else
        Exit Sub
    End If


    Select Case Me.ActiveControl.Name
    
        Case "docViewer"
            
            'EnableFolderShortcut False
            EnableFileShortcut True
            
            If docViewer.GetSelectedTabIndex() = 2 Then
                'mnuCheckIn.Enabled = False
                'mnuCheckOut.Enabled = False
                mnuNew.Enabled = False
                mnuImportDoc.Enabled = False
                mnuRefresh.Enabled = False
                mnuBatchUpdate.Enabled = False
                mnuDuplicate.Enabled = False
                mnuView.Enabled = False
                mnuMarkSelAs.Enabled = False
                
                tbShortcut.Buttons("new").Enabled = False
                tbShortcut.Buttons("import").Enabled = False
                tbShortcut.Buttons("view").Enabled = False
                tbShortcut.Buttons("refresh").Enabled = False
                tbShortcut.Buttons("duplicate").Enabled = False
                tbShortcut.Buttons("checkin").Enabled = False
                
            Else
                'mnuCheckIn.Enabled = True
                'mnuCheckOut.Enabled = True
                mnuNew.Enabled = True
                mnuImportDoc.Enabled = True
                mnuRefresh.Enabled = True
                mnuBatchUpdate.Enabled = True
                mnuDuplicate.Enabled = True
                mnuView.Enabled = True
                mnuMarkSelAs.Enabled = True
            End If
            
            docViewer.HaveConfidentialAccess = AppCurrentUser.bCanViewConfidential

            
        Case "trvFolderList"
            
            'EnableFolderShortcut True
            'EnableFileShortcut False

    
    End Select
End Sub

Private Sub EnableFileShortcut(ByVal enable As Boolean)
    tbShortcut.Buttons(7).Enabled = enable
    tbShortcut.Buttons(8).Enabled = enable
    tbShortcut.Buttons(9).Enabled = enable
    tbShortcut.Buttons(10).Enabled = enable
    tbShortcut.Buttons(11).Enabled = enable
    tbShortcut.Buttons(12).Enabled = enable
    tbShortcut.Buttons(13).Enabled = enable
    tbShortcut.Buttons(14).Enabled = enable
    tbShortcut.Buttons(15).Enabled = enable
    tbShortcut.Buttons(16).Enabled = enable
    
    tbShortcut.Buttons(18).Enabled = enable
End Sub

Private Sub EnableFolderShortcut(ByVal enable As Boolean)
    tbShortcut.Buttons(2).Enabled = enable
    tbShortcut.Buttons(3).Enabled = enable
    tbShortcut.Buttons(4).Enabled = enable
    tbShortcut.Buttons(5).Enabled = enable
End Sub

Private Sub tmrResize_Timer()
    On Error Resume Next
    GetCursorPos cursor_pos
    picLeftContainer.Width = (((cursor_pos.x * Screen.TwipsPerPixelX) - Me.Left)) - 90
    
    PositionDocViewer
End Sub

Private Sub PositionDocViewer()
    On Error Resume Next
    
    docViewer.Top = picLeftContainer.Top
    docViewer.Left = picSeparator.Left + picSeparator.Width
    
    docViewer.Width = Me.ScaleWidth - docViewer.Left
    docViewer.Height = Me.ScaleHeight - (docViewer.Top + statBar.Height)
End Sub


Private Sub trvFolderList_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Node.EnsureVisible
    
    statBar.Panels(1).Text = Node.FullPath
    
    CheckFolderLeverRestrictions Replace(Node.Tag, "ID:", "")
    docViewer.folderId = Replace(Node.Tag, "ID:", "")
End Sub

Private Sub trvFolderList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
            
        PopupMenu mnuFolderAction
    End If
End Sub


Private Sub LoadDirectory(ByVal DirectoryName As String, Optional ParentKey As String, Optional ForceToExpand As Boolean, Optional IsNewlyInserted As Boolean)
    DoEvents
    
    On Error GoTo err
    
    ShowWaiting Me, WaitingDisplay
    
    
    Dim rsFolder As recordset
    
    If ParentKey = "" Then
        Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
    Else
    
        If CurrentFolderAccess.folderId = Replace(ParentKey, "ID:", "") And CurrentFolderAccess.bDenyFolderAccess = True Then
            PrompAccessDeniedForFolder
            Exit Sub
        End If
    
        If IsNewlyInserted = True Then
            Set rsFolder = GetRecords("SELECT TOP 1 * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [ID] DESC,[DirectoryName],[FolderName] ASC")
        Else
            Set rsFolder = GetRecords("SELECT * FROM vw_Folders WHERE ParentFolderID=" & Replace(ParentKey, "ID:", "") & " AND DirectoryName='" & DirectoryName & "' ORDER BY [DirectoryName],[FolderName] ASC")
        End If
    End If
    
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
                    
                    statBar.Panels(1).Text = "General"
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
    
    HideWaiting Me, WaitingDisplay 'frmWaiting.Terminate
    Exit Sub
err:
    If err.Number = 35602 Or err.Number = 91 Or err.Number = 35605 Then
        Resume Next
    Else
        HideWaiting Me, WaitingDisplay 'frmWaiting.Terminate
        'InputBox err.Description, "", err.Number
    End If
End Sub


Private Sub trvFolderList_NodeClick(ByVal Node As MSComctlLib.Node)
    statBar.Panels(1).Text = Node.FullPath
    
    CheckFolderLeverRestrictions Replace(Node.Tag, "ID:", "")
    docViewer.folderId = Replace(Node.Tag, "ID:", "")
End Sub

Private Sub trvFolderList_Expand(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Node.EnsureVisible
    CheckFolderLeverRestrictions Replace(Node.Tag, "ID:", "")
    docViewer.folderId = Replace(Node.Tag, "ID:", "")
    statBar.Panels(1).Text = Node.FullPath
    If Node.children > 0 Then
        If Node.Child.Tag = "loading" Then
            'Remove the temporary child
            trvFolderList.Nodes.Remove (Node.Child.Index)

            LoadDirectory Node.FullPath & "\", Node.Tag
        End If
    End If
End Sub


Private Sub trvFolderList_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo err
    DoEvents
    
    Dim strRawItem As String
    Dim strItems() As String
    Dim strFolderID As String
    Dim targetNode As Node
    Dim currentNode As Node
    Dim i As Integer
    
    Set targetNode = trvFolderList.HitTest(x, y)
    If targetNode Is Nothing Then Exit Sub
    
    Set trvFolderList.DropHighlight = Nothing
    
    ShowWaiting Me, WaitingDisplay
    
    
    'Move the selected item(s)
    If InStr(1, data.GetData(1), "DROP_ITEM") = 1 Then
        If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
        strRawItem = Replace(data.GetData(1), "DROP_ITEM", "")
        strItems = Split(strRawItem, "~")
        strFolderID = strItems(0)
        
        If strFolderID <> Replace(targetNode.Tag, "ID:", "") Then
            For i = 1 To UBound(strItems)
                If ExecSQL("UPDATE tbl_Files SET [FolderID]=" & Replace(targetNode.Tag, "ID:", "") & " WHERE [ID]=" & strItems(i)) = False Then
                    HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
                    MsgBox "Error occur in moving the document.", vbCritical
                    Exit Sub
                End If
            Next i
            
            HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
            RefreshDocList
        End If
        Exit Sub
    End If
    
    'Move the selected folder
    Set currentNode = trvFolderList.SelectedItem
    If InStr(1, data.GetData(1), "DROP_FOLDER") = 1 And mnuEdit.Enabled = True Then
        If AppCurrentUser.bCanEditFolder = False Then PrompAccessDenied: Exit Sub
        strRawItem = Replace(data.GetData(1), "DROP_FOLDER", "")
        strFolderID = strRawItem
        
        If strFolderID <> Replace(targetNode.Tag, "ID:", "") Then
            If ExecSQL("UPDATE tbl_Folders SET [ParentFolderID]=" & Replace(targetNode.Tag, "ID:", "") & ",[DirectoryName]='" & targetNode.FullPath & "\'  WHERE [ID]=" & strFolderID) = False Then
                HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
                MsgBox "Error occur in moving the document.", vbCritical
                Exit Sub
            Else
                UpdateSubDir strFolderID, targetNode.FullPath & "\" & currentNode.Text & "\"
            End If
            
            Dim tempNode As Node
            
            RemoveSubNode currentNode
            trvFolderList.Nodes.Remove (currentNode.Index)
            Set tempNode = trvFolderList.Nodes.Add(targetNode.Tag, tvwChild, currentNode.Tag, currentNode.Text, 2, 1)
            tempNode.Tag = currentNode.Tag
            tempNode.Selected = True
            tempNode.EnsureVisible
            tempNode.Expanded = True
            
            HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
            LoadDirectory tempNode.FullPath & "\", tempNode.Tag
            

        End If
    End If
    
    HideWaiting Me, WaitingDisplay 'frmWaiting.Terminate
    
    Exit Sub
err:
    HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
End Sub

Private Sub RemoveSubNode(ByRef srcNode As Node)
    If srcNode.children > 0 Then
        Do While srcNode.children > 0
            trvFolderList.Nodes.Remove (srcNode.Child.FirstSibling.Index)
        Loop
    End If
End Sub

Private Sub trvFolderList_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Dim selNode As Node
    Set selNode = trvFolderList.HitTest(x, y)
    trvFolderList.DropHighlight = selNode
End Sub

Private Sub trvFolderList_OLESetData(data As MSComctlLib.DataObject, DataFormat As Integer)
    DoEvents
    
    If trvFolderList.SelectedItem.Text <> trvFolderList.SelectedItem.FullPath Then data.SetData "DROP_FOLDER" & Replace(trvFolderList.SelectedItem.Tag, "ID:", "")
End Sub


Public Sub CheckFolderLeverRestrictions(ByVal folderId As String)
    
    On Error Resume Next
    
    ShowWaiting Me, WaitingDisplay
    
    
    Dim rsRecord As recordset
    
    'Set the default
    With CurrentFolderAccess
        .bDenyFolderAccess = False
        .bDenyFolderEdit = False
        .bDenyFolderDelete = False
        .bDenyOpenFile = False
        .bDenyCreateFile = False
        .bDenyEditFile = False
        .bDenyDeleteFile = False
        .bDenyCheckOut = False
        .bDenyFileImport = False
        .bDenyFileExport = False
        .folderId = Val(folderId)
    End With
            
    If AppCurrentUser.bIsSysAdmin = False Then
        Set rsRecord = GetRecords("SELECT * FROM vw_FolderRestrictionSet WHERE [FolderID]=" & folderId & " AND [UserGroupID]=" & AppCurrentUser.UserGroupId)
        If Not rsRecord Is Nothing Then
            If rsRecord.RecordCount > 0 Then
                With CurrentFolderAccess
                    If Val(rsRecord.Fields("DenyFolderAccess")) <> 0 Then .bDenyFolderAccess = True
                    If Val(rsRecord.Fields("DenyFolderEdit")) <> 0 Then .bDenyFolderEdit = True
                    If Val(rsRecord.Fields("DenyFolderDelete")) <> 0 Then .bDenyFolderDelete = True
                    If Val(rsRecord.Fields("DenyOpenFile")) <> 0 Then .bDenyOpenFile = True
                    If Val(rsRecord.Fields("DenyCreateFile")) <> 0 Then .bDenyCreateFile = True
                    If Val(rsRecord.Fields("DenyEditFile")) <> 0 Then .bDenyEditFile = True
                    If Val(rsRecord.Fields("DenyDeleteFile")) <> 0 Then .bDenyDeleteFile = True
                    If Val(rsRecord.Fields("DenyCheckOut")) <> 0 Then .bDenyCheckOut = True
                    If Val(rsRecord.Fields("DenyFileImport")) <> 0 Then .bDenyFileImport = True
                    If Val(rsRecord.Fields("DenyFileExport")) <> 0 Then .bDenyFileExport = True
                    .folderId = Val(folderId)
                End With
            End If
        End If
        Set rsRecord = Nothing
    End If
        
    HideWaiting Me, WaitingDisplay 'frmWaiting.ForceTerminate
    
End Sub
