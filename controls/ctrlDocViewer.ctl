VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ctrlDocViewer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8220
   ScaleWidth      =   12030
   Begin MSComctlLib.ImageList imgTab 
      Left            =   10215
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctrlDocViewer.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctrlDocViewer.ctx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctrlDocViewer.ctx":06A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox panelDesc 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   2115
      Left            =   105
      ScaleHeight     =   2115
      ScaleWidth      =   10830
      TabIndex        =   17
      Top             =   5460
      Width           =   10830
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         Height          =   1935
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   225
         Width           =   10860
      End
      Begin VB.Label lblFileDesc 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "  File Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   225
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2955
      End
   End
   Begin MSComDlg.CommonDialog dlgFileImport 
      Left            =   405
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files|*.*"
      Flags           =   1
   End
   Begin VB.PictureBox picExplorerContainer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7440
      Left            =   735
      ScaleHeight     =   7380
      ScaleWidth      =   11430
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   11490
      Begin VB.PictureBox picTemp 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   9555
         ScaleHeight     =   360
         ScaleWidth      =   1560
         TabIndex        =   15
         Top             =   5775
         Visible         =   0   'False
         Width           =   1560
      End
      Begin MSComctlLib.ImageList imgDocIco16 
         Left            =   8775
         Top             =   4230
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
               Picture         =   "ctrlDocViewer.ctx":0A42
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lstvSearchResult 
         Height          =   4470
         Left            =   180
         TabIndex        =   14
         Top             =   2070
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   7885
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "imgDocIco32"
         SmallIcons      =   "imgDocIco16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Location"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Original Author"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Doc. No"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Doc. Index"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Doc. Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Doc. Type"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Physical Location"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "File Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Alert Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Alert Note"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Date Created"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Last Modified"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Created By"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Last Modified By"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Is Confidential"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Check Out By"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Date Check Out"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "CheckOutByPK"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.ComboBox cmbFields 
         Height          =   315
         ItemData        =   "ctrlDocViewer.ctx":0DDC
         Left            =   1350
         List            =   "ctrlDocViewer.ctx":0E19
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   495
         Width           =   6270
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   1050
         Left            =   7740
         Picture         =   "ctrlDocViewer.ctx":0F1F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Search"
         Top             =   495
         Width           =   960
      End
      Begin VB.ComboBox cbxCriteria 
         Height          =   315
         ItemData        =   "ctrlDocViewer.ctx":1786
         Left            =   1350
         List            =   "ctrlDocViewer.ctx":17A2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   855
         Width           =   2220
      End
      Begin VB.TextBox txtFindWhat 
         Height          =   330
         Left            =   1350
         TabIndex        =   2
         Top             =   1215
         Width           =   6270
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   11
         Top             =   1215
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   40304643
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Index           =   1
         Left            =   3630
         TabIndex        =   12
         Top             =   1215
         Visible         =   0   'False
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   40304643
         CurrentDate     =   38207
      End
      Begin MSComctlLib.ImageList imgDocIco24 
         Left            =   9450
         Top             =   4230
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
               Picture         =   "ctrlDocViewer.ctx":182A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgDocIco32 
         Left            =   10080
         Top             =   4200
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
               Picture         =   "ctrlDocViewer.ctx":1FA4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   3150
         TabIndex        =   13
         Top             =   1245
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Search Result(s) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   1845
         Width           =   2805
      End
      Begin VB.Label Label4 
         Caption         =   "Enter the criteria bellow :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   225
         Width           =   2805
      End
      Begin VB.Label Label3 
         Caption         =   "Look In:"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Criteria:"
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView lstvExplorer 
      Height          =   7215
      Left            =   100
      TabIndex        =   6
      Top             =   450
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   12726
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "imgDocIco32"
      SmallIcons      =   "imgDocIco16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   20
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
         Text            =   "Alert Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Alert Note"
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
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Check Out By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Date Check Out"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "CheckOutByPK"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabDocViewer 
      Height          =   8115
      Left            =   45
      TabIndex        =   4
      Top             =   315
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   14314
      ImageList       =   "imgTab"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Document(s)"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
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
   Begin MSComDlg.CommonDialog dlgCheckIn 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   1
   End
   Begin VB.Label lblFileList 
      BackColor       =   &H8000000D&
      Caption         =   " File List"
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
      TabIndex        =   16
      Top             =   0
      Width           =   6210
   End
End
Attribute VB_Name = "ctrlDocViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Const m_def_HaveChanges = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_HaveChanges As Boolean
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

Dim m_FolderId As Long
Dim m_SearchSQL As String
Dim m_HaveConfidentialAccess As Boolean
Dim iCurrentExCol As Integer
Dim iCurrentSrchCol As Integer
Dim bIsExplorerMouseDwn As Boolean
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstvExplorer,lstvExplorer,-1,MouseDown
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Event NavChange(NaveInfo As String)


Private Sub cbxCriteria_Click()

    If cbxCriteria.ListIndex = 7 Then
        dtpDate(0).Visible = True
        dtpDate(1).Visible = True
        txtFindWhat.Visible = False
    Else
        txtFindWhat.Visible = True
        dtpDate(0).Visible = False
        dtpDate(1).Visible = False
    End If
    
End Sub

Private Sub cmdSearch_Click()
    m_SearchSQL = ""
    
    
    On Error GoTo err
    Dim strCondition As String
    Dim strConditionForNull As String
    'Initialize the fields
    strCondition = Replace(cmbFields.Text, "/", "") 'ex. City/Town for tblCustomer
    strCondition = Replace(cmbFields.Text, " ", "")
    strCondition = "[" & strCondition & "]"
    strConditionForNull = strCondition & " Is Null"
    
    'Initialize the operation used
    'First operation
    Select Case cbxCriteria.ListIndex
        Case 0: strCondition = strCondition & " LIKE '%" & txtFindWhat.Text & "%'"
        Case 1: strCondition = strCondition & " = '" & txtFindWhat.Text & "'"
        Case 2: strCondition = strCondition & " <> '" & txtFindWhat.Text & "'"
        Case 3: strCondition = strCondition & " > '" & txtFindWhat.Text & "'"
        Case 4: strCondition = strCondition & " >= '" & txtFindWhat.Text & "'"
        Case 5: strCondition = strCondition & " < '" & txtFindWhat.Text & "'"
        Case 6: strCondition = strCondition & " <= '" & txtFindWhat.Text & "'"
        Case 7
            If AppDBType = adDBTypeSQLServer Then
                strCondition = strCondition & " BETWEEN '" & Format$(dtpDate(0).Value, "yyyy/MM/dd") & "' AND '" & Format$(dtpDate(1).Value, "yyyy/MM/dd") & "'"
            ElseIf AppDBType = adDBTypeMSAccess Then
                strCondition = strCondition & " BETWEEN #" & Format$(dtpDate(0).Value, "yyyy/MM/dd") & "# AND #" & Format$(dtpDate(1).Value, "yyyy/MM/dd") & "#"
            End If
            
    End Select
    
    If cbxCriteria.ListIndex <> 0 And cbxCriteria.ListIndex <> 7 Then If IsNumeric(txtFindWhat.Text) = True Then strCondition = Replace(strCondition, "'", "")
    If cbxCriteria.ListIndex <> 7 Then If txtFindWhat.Text = "" Then strCondition = strConditionForNull
     
    If AppCurrentUser.bCanManageTemplates = False Then
        strCondition = "((" & strCondition & ") AND ([Location] NOT LIKE '%Templates\'))"
    End If
    m_SearchSQL = strCondition
    
    LoadFilesViaSQL
    'Clear used variables
    strCondition = vbNullString


    Exit Sub
err:

End Sub

Private Sub lstvExplorer_Click()
    If lstvExplorer.ListItems.Count <> 0 Then
        RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
        txtDesc.Text = lstvExplorer.SelectedItem.ToolTipText
    End If
End Sub

Private Sub lstvExplorer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sort the listview
    If ColumnHeader.Index - 1 <> iCurrentExCol Then
        lstvExplorer.SortOrder = 0
    Else
        lstvExplorer.SortOrder = Abs(lstvExplorer.SortOrder - 1)
    End If
    lstvExplorer.SortKey = ColumnHeader.Index - 1
    
    lstvExplorer.Sorted = True
    iCurrentExCol = ColumnHeader.Index - 1
End Sub

Private Sub lstvExplorer_DblClick()
    OpenFile
End Sub

Private Sub lstvExplorer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lstvExplorer_Click
End Sub

Private Sub lstvExplorer_LostFocus()
    If lstvExplorer.ListItems.Count > 0 Then
        If lstvExplorer.SelectedItem Is Nothing Then
            Set lstvExplorer.SelectedItem = lstvExplorer.ListItems(1)
            lstvExplorer.SelectedItem.Selected = True
        End If
    End If
End Sub


Private Sub lstvExplorer_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEvents
    
    On Error GoTo err:
    Dim i As Integer
    If data.Files.Count > 0 Then
        For i = 1 To data.Files.Count
            PerformImport data.Files.item(i)
        Next i
    End If
    
    Exit Sub
err:
    'MsgBox "Error occur while reading file(s).", vbCritical
End Sub

Private Sub lstvExplorer_OLESetData(data As MSComctlLib.DataObject, DataFormat As Integer)
    DoEvents
    
    Dim item As ListItem
    Dim selItems As String

    For Each item In lstvExplorer.ListItems
        If item.Selected = True Then
            If selItems = "" Then
                selItems = item.Tag
            Else
                selItems = selItems & "~" & item.Tag
            End If
        End If
    Next
    
    data.SetData "DROP_ITEM" & folderId & "~" & selItems
    
End Sub

Private Sub lstvSearchResult_Click()
    If lstvSearchResult.ListItems.Count <> 0 Then
        RaiseEvent NavChange("Record: " & lstvSearchResult.SelectedItem.Index & " of " & lstvSearchResult.ListItems.Count)
    End If
End Sub

Private Sub lstvSearchResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sort the listview
    If ColumnHeader.Index - 1 <> iCurrentSrchCol Then
        lstvSearchResult.SortOrder = 0
    Else
        lstvSearchResult.SortOrder = Abs(lstvSearchResult.SortOrder - 1)
    End If
    lstvSearchResult.SortKey = ColumnHeader.Index - 1
    
    lstvSearchResult.Sorted = True
    iCurrentSrchCol = ColumnHeader.Index - 1
End Sub

Private Sub lstvSearchResult_DblClick()
    OpenFile
End Sub

Private Sub lstvSearchResult_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lstvSearchResult_Click
End Sub

Private Sub lstvSearchResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub picExplorerContainer_Resize()
    On Error Resume Next

    lstvSearchResult.Width = picExplorerContainer.Width - 360
    lstvSearchResult.Height = picExplorerContainer.Height - (lstvSearchResult.Top + 180)

End Sub

Private Sub tabDocViewer_Click()
    If tabDocViewer.SelectedItem.Index = 1 Then
        lstvExplorer.Visible = True
        txtDesc.Visible = True
        picExplorerContainer.Visible = False
        
        lstvExplorer.ZOrder
        LoadFiles
'        If lstvExplorer.ListItems.Count = 0 Then
'            RaiseEvent NavChange("Record: 0 of 0")
'        Else
'            RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
'        End If
    ElseIf tabDocViewer.SelectedItem.Index = 2 Then
        lstvExplorer.Visible = False
        txtDesc.Visible = False
        picExplorerContainer.Visible = True
        
        picExplorerContainer.ZOrder
        LoadFilesViaSQL
'        If lstvSearchResult.ListItems.Count = 0 Then
'            RaiseEvent NavChange("Result(s): 0 of 0")
'        Else
'            RaiseEvent NavChange("Result(s): " & lstvSearchResult.SelectedItem.Index & " of " & lstvSearchResult.ListItems.Count)
'        End If
    End If
End Sub

Private Sub txtFindWhat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSearch_Click
End Sub

Private Sub UserControl_Initialize()
    cbxCriteria.ListIndex = 0
    cmbFields.ListIndex = 0
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    lblFileList.Width = UserControl.ScaleWidth
    
    tabDocViewer.Width = UserControl.ScaleWidth - 90
    tabDocViewer.Height = UserControl.ScaleHeight - (90 + lblFileList.Height)
    
    lstvExplorer.Top = tabDocViewer.Top + 200 + lblFileList.Height
    lstvExplorer.Width = UserControl.ScaleWidth - 210
    lstvExplorer.Height = UserControl.ScaleHeight - (lstvExplorer.Top + panelDesc.Height + 105)
    
    panelDesc.Top = lstvExplorer.Top + lstvExplorer.Height + 50
    panelDesc.Height = 2015 + txtDesc.Top 'tabDocViewer.Height - (lstvExplorer.Height + 550)
    
    txtDesc.Height = 2015
    panelDesc.Width = lstvExplorer.Width
    lblFileDesc.Width = lstvExplorer.Width
    txtDesc.Width = lstvExplorer.Width
    
    
    picExplorerContainer.Top = lstvExplorer.Top
    picExplorerContainer.Left = lstvExplorer.Left
    
    picExplorerContainer.Width = lstvExplorer.Width
    picExplorerContainer.Height = lstvExplorer.Height + panelDesc.Height + 75
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub
'
'Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseDown(Button, Shift, X, Y)
'End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_HaveChanges = m_def_HaveChanges
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
End Sub

Private Sub lstvExplorer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Public Property Get folderId() As Long
    folderId = m_FolderId
End Property

Public Property Let folderId(ByVal New_FolderId As Long)
    If m_FolderId = New_FolderId Then Exit Property
    m_FolderId = New_FolderId
    PropertyChanged "FolderId"
    
    If CurrentFolderAccess.folderId = New_FolderId And CurrentFolderAccess.bDenyFolderAccess = True Then
        PrompAccessDeniedForFolder
        Exit Property
    End If
    
    LoadFiles
End Property

Public Sub LoadFiles()
    DoEvents
    
    lstvExplorer.ListItems.Clear
    txtDesc.Text = ""
    
    Dim rsDocFiles As recordset
    Set rsDocFiles = GetRecords("SELECT [FileName],[Title],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[FileSize],[AlertDate],[AlertNote],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy],[IsConfidential],[Status],[CheckOutBy],[DateCheckOut],[CheckOutByFK],[ID],[Description] FROM vw_FileInfoOnly WHERE [FolderID]=" & m_FolderId & "")
    
    
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay

        
    If Not rsDocFiles Is Nothing Then
        If rsDocFiles.RecordCount > 0 Then
            FillListView lstvExplorer, rsDocFiles, 20, -1, False, True, "ID", "Description", "IsConfidential", "FileName", imgDocIco32, imgDocIco16, picTemp
            RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
        Else
            RaiseEvent NavChange("Record: 0 of 0")
        End If

        m_HaveChanges = True
        Set rsDocFiles = Nothing
    End If
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
End Sub

Public Sub LoadFilesViaSQL()
    If m_SearchSQL = "" Then Exit Sub

    DoEvents
    
    On Error GoTo err
    lstvSearchResult.ListItems.Clear
    txtDesc.Text = ""
    
    Dim rsDocFiles As recordset
    Set rsDocFiles = GetRecords("SELECT [FileName],[Title],[Location],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[FileSize],[AlertDate],[AlertNote],[DateCreated],[LastModified],[CreatedBy],[LastModifiedBy],[IsConfidential],[Status],[CheckOutBy],[DateCheckOut],[CheckOutByFK],[ID],[Description] FROM vw_FileInfoOnly WHERE " & m_SearchSQL & "")
    
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
        
    If Not rsDocFiles Is Nothing Then
        
        If rsDocFiles.RecordCount > 0 Then
            FillListView lstvSearchResult, rsDocFiles, 21, -1, False, True, "ID", "Description", "IsConfidential", "FileName", imgDocIco32, imgDocIco16, picTemp
            RaiseEvent NavChange("Result(s): " & lstvSearchResult.SelectedItem.Index & " of " & lstvSearchResult.ListItems.Count)
        Else
            RaiseEvent NavChange("Result(s): 0 of 0")
        End If

        m_HaveChanges = True
        Set rsDocFiles = Nothing
    End If
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    Exit Sub
err:

    If err.Number = -2147352571 Then
        MsgBox "Invalid search operation.", vbExclamation
        HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    ElseIf err.Number = 3001 Then
        Resume Next
    End If
End Sub

Public Sub ChangeView(ByVal ViewCode As Integer)
    lstvExplorer.View = ViewCode - 1
End Sub

Public Sub ProccessCommand(ByVal CmdCode As String)
    Select Case CmdCode
    
        Case "OPEN_DOC": OpenFile
        Case "OPEN_WITH": OpenFile True
        Case "NEW_DOC": CreateNewDoc
        Case "EDIT_DOC": EditDoc
        Case "IMPORT_DOC": ImportDoc
        Case "DELETE_DOC": DeleteDoc
        Case "REFRESH_DOC"
            If tabDocViewer.SelectedItem.Index = 1 Then
                LoadFiles
            Else
                LoadFilesViaSQL
            End If
        Case "SELECT_ALL": SelectAll
        Case "MARK_CONFI": MarkAsConfidentialAccess True
        Case "MARK_NONCONFI": MarkAsConfidentialAccess False
        Case "EXPORT_DOC": ExportDoc
        Case "UPDATE_BATCH": UpdateBatch
        Case "DUPLICATE_DOC": DuplicateDoc
        
            
        
    End Select

End Sub


Private Sub OpenFile(Optional isOpenWith As Boolean)
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyOpenFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    If HaveFolderLevelAccess(GetSrcFileID(), "open") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    Dim recCount As Long
    
    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        LunchSelectedFile Val(lstvExplorer.SelectedItem.Tag), isOpenWith
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        LunchSelectedFile Val(lstvSearchResult.SelectedItem.Tag), isOpenWith
    End If

End Sub

Private Sub LunchSelectedFile(ByVal lFileId As Long, Optional isOpenWith As Boolean)

    Dim strDocName As String
    Dim strDocFileLocation As String
    Dim rsDocument As recordset
    
    Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & lFileId & "")
            
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
    If Not rsDocument Is Nothing Then
        If rsDocument.RecordCount > 0 Then
        
            strDocName = rsDocument.Fields("FileName")
            strDocFileLocation = App.Path & "\temp\" & GetFileNameWithOutExt(strDocName) & "(READ ONLY)" & "_" & GenerateFileName("." & GetFileExt(strDocName))
        
            If DownloadFileFromDB(strDocFileLocation, rsDocument, "FileObj") = False Then
                MsgBox "Error occur while downloading the document.", vbCritical
                Unload Me
                Exit Sub
            End If
            
            SetAttr strDocFileLocation, vbReadOnly
            
            If isOpenWith = True Then
                LunchFileWithDialog strDocFileLocation
            Else
                OpenURL strDocFileLocation, frmFileManager.hwnd
            End If
            
        End If
    End If
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
    Set rsDocument = Nothing
    strDocName = vbNullString
    strDocFileLocation = vbNullString

End Sub


Private Sub CreateNewDoc()
    
    If AppCurrentUser.bCanAdd = False Then PrompAccessDenied: Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyCreateFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    
    frmTemplateExplorer.Show vbModal, frmFileManager
    If LastUseFileNamePath = "" Then Exit Sub
    
    If lstvExplorer.ListItems.Count > 0 Then lstvExplorer.SelectedItem.Selected = False
    PerformImport LastUseFileNamePath, LastUseFileId
    
    frmUpdateFile.lFileId = Val(lstvExplorer.SelectedItem.Tag)
    frmUpdateFile.lFolderId = m_FolderId
    frmUpdateFile.IsEditMode = False
    frmUpdateFile.Show vbModal, frmFileManager
    
End Sub

Private Sub EditDoc()
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyEditFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    If HaveFolderLevelAccess(GetSrcFileID(), "edit") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    Dim recCount As Long
    Dim strDocName As String
    Dim strDocId As Long

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        strDocName = lstvExplorer.SelectedItem.Text
        strDocId = Val(lstvExplorer.SelectedItem.Tag)
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        strDocName = lstvSearchResult.SelectedItem.Text
        strDocId = Val(lstvSearchResult.SelectedItem.Tag)
    End If
    
   

    frmUpdateFile.lFileId = strDocId
    frmUpdateFile.lFolderId = m_FolderId
    frmUpdateFile.IsEditMode = True
    frmUpdateFile.Show vbModal, frmFileManager
    
End Sub

Private Sub ImportDoc()
    If AppCurrentUser.bCanImport = False Then PrompAccessDenied: Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyFileImport = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    On Error GoTo cError
    
    Dim i As Integer
    Dim myFiles() As String
    Dim myPath As String
    
    Const CD_FLAGS = cdlOFNAllowMultiselect + cdlOFNExplorer + cdlOFNLongNames
    
    With dlgFileImport
        .MaxFileSize = 32000 'this will max out the buffer for the filenames array for large selections. *NEW*
        .DialogTitle = "Import File(s)"
        .fileName = ""
        .Flags = CD_FLAGS 'this is where we tell it to use multiselect
        .ShowOpen
        
        If .fileName = "" Then
            Exit Sub
        Else
            DoEvents
            
            SelectNone
            
            myFiles = Split(.fileName, vbNullChar) 'the Filename returned is delimeted by a null character because we selected the cdlOFNLongNames flag
            
            ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
            
            
            Select Case UBound(myFiles)
                Case 0 'if only one was selected we are done
                    PerformImport myFiles(0)
                Case Is > 0 'if more than one, we need to loop through it and append the root directory
                    If Not lstvExplorer.SelectedItem Is Nothing Then lstvExplorer.SelectedItem.Selected = False
                    For i = 1 To UBound(myFiles)
                        myPath = myFiles(0) & IIf(Right(myFiles(0), 1) <> "\", "\", "") & myFiles(i)
                        PerformImport myPath, , "Imported file."
                    Next i
            End Select
            
            HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
             
        End If

    End With
    Exit Sub

cError:
    MsgBox err.Description, vbCritical, "Unexpected Error"
       
End Sub

Private Sub PerformImport(ByVal strDocLocation As String, Optional removeUnwantedFileString As String, Optional importDesc As String)
    
    If strDocLocation <> "" And InStr(1, strDocLocation, ".") > 0 Then
    
        Dim rsImportDoc As recordset
        Dim strDocFileName As String
        Dim lngNoOfDuplicate As Long
        
        Set rsImportDoc = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=-1")
        
        strDocFileName = GetFileNameFromPath(strDocLocation)
        strDocFileName = Replace(strDocFileName, "_" & removeUnwantedFileString, "")
        
        lngNoOfDuplicate = GetRecordCount("SELECT [ID] FROM tbl_Files WHERE (([FolderID]=" & m_FolderId & ") AND ([FileName]='" & strDocFileName & "') )")
        
        
        If lngNoOfDuplicate <> 0 Then
            Dim i As Integer
            Do
                i = i + 1
                lngNoOfDuplicate = GetRecordCount("SELECT [ID] FROM tbl_Files WHERE (([FolderID]=" & m_FolderId & ") AND ([FileName]='" & GetNameFromFileName(strDocFileName) & "_" & i & "." & GetFileExt(strDocFileName) & "') )")
            Loop While (lngNoOfDuplicate > 0)
            strDocFileName = GetNameFromFileName(strDocFileName) & "_" & i & "." & GetFileExt(strDocFileName)
        End If
    
        If Not rsImportDoc Is Nothing Then
            
            With rsImportDoc
                
                .AddNew
                
                .Fields("FileName") = strDocFileName
                
                SaveFileToDB strDocLocation, rsImportDoc, "FileObj"

                .Fields("FileSize") = GetFileSizeString(strDocLocation)
                .Fields("FileSizeNo") = GetFileSize(strDocLocation)
                .Fields("Description") = importDesc
                .Fields("FolderID") = m_FolderId
                .Fields("DateCreated") = Now
                .Fields("CreatedBy") = AppCurrentUser.CompleteName
                
                '.Update
                
            End With
            
            If SaveRecord("SELECT [FileName],[Description],[FileSize],[FileSizeNo],[FileObj],[FolderID],[DateCreated],[CreatedBy] FROM tbl_Files " & _
                          "WHERE [ID]=-1", rsImportDoc, , , "tbl_Files") = 1 Then
                
                
                Dim lstItem As ListItem
                
                Set lstItem = lstvExplorer.ListItems.Add(, , strDocFileName)
                lstItem.Icon = ExtractIcon(strDocFileName, imgDocIco32, picTemp, 32)
                lstItem.SmallIcon = ExtractIcon(strDocFileName, imgDocIco16, picTemp, 16)
                lstItem.Tag = LAST_GENERATED_IDENTITY
                
                lstItem.SubItems(8) = rsImportDoc.Fields("FileSize")
                lstItem.SubItems(11) = FormatRecord(rsImportDoc.Fields("DateCreated"))
                lstItem.SubItems(13) = rsImportDoc.Fields("CreatedBy")
                lstItem.SubItems(15) = "No"
                
                lstItem.ToolTipText = rsImportDoc.Fields("Description")
                lstItem.Selected = True
                lstItem.EnsureVisible
                
                Set lstItem = Nothing
                
                RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
                m_HaveChanges = True
                lstvExplorer.SetFocus
                
            End If
        
            Set rsImportDoc = Nothing
        End If
    End If
        
    
End Sub

Private Sub DeleteDoc()
    If AppCurrentUser.bCanDelete = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyDeleteFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    If HaveFolderLevelAccess(GetSrcFileID(), "delete") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    Dim recCount As Long
    

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
    End If
    
    If MsgBox("Are you sure you want to delete the selected file(s)?", vbCritical + vbYesNo, "Confirm File Deletion") = vbYes Then
        Dim item As ListItem
    
        ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
        
            
        If tabDocViewer.SelectedItem.Index = 1 Then
            For Each item In lstvExplorer.ListItems
                If item.Selected = True Then
                     If item.SubItems(18) = "" Then
                        If item.SubItems(15) = "Yes" Then
                            If m_HaveConfidentialAccess = True Then If DeleteDocOneByOne(Val(item.Tag)) = False Then Exit For
                        Else
                            If DeleteDocOneByOne(Val(item.Tag)) = False Then Exit For
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            Next
            
            LoadFiles
        Else
            For Each item In lstvSearchResult.ListItems
                If item.Selected = True Then
                    If item.SubItems(19) = "" Then
                        If item.SubItems(16) = "Yes" Then
                            If m_HaveConfidentialAccess = True Then If DeleteDocOneByOne(Val(item.Tag)) = False Then Exit For
                        Else
                            If DeleteDocOneByOne(Val(item.Tag)) = False Then Exit For
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            Next
            
            LoadFilesViaSQL
        End If
        
        HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
        
        m_HaveChanges = True

    End If
    
End Sub

Private Function DeleteDocOneByOne(ByVal strDocId As Long) As Boolean
       
    Dim rsDelFile As recordset
    
    Set rsDelFile = GetRecords("SELECT [IsDeleted],[DeletedBy],[DeletionDate] FROM tbl_Files WHERE [ID]=" & strDocId)
    
    rsDelFile.Fields("IsDeleted") = 1
    rsDelFile.Fields("DeletedBy") = AppCurrentUser.CompleteName
    rsDelFile.Fields("DeletionDate") = Now
    
    If SaveRecord("", rsDelFile, , True) = 1 Then
        DeleteDocOneByOne = True
    End If
    Set rsDelFile = Nothing
    
    'DeleteDocOneByOne = ExecSQL("UPDATE tbl_Files SET [IsDeleted]=1,[DeletedBy]='" & AppCurrentUser.CompleteName & "',[DeletionDate]='" & Now & "' WHERE [ID]=" & strDocId & "")

End Function


Public Function GetCurrentRecordCount() As Long
    If tabDocViewer.SelectedItem.Index = 1 Then
        GetCurrentRecordCount = lstvExplorer.ListItems.Count
    Else
        GetCurrentRecordCount = lstvSearchResult.ListItems.Count
    End If
End Function

Public Function GetSelectedTabIndex() As Integer
    GetSelectedTabIndex = tabDocViewer.SelectedItem.Index
End Function


Public Sub UpdateListForNewDoc(ByRef srcRecords As recordset)
    
    Dim lstItem As ListItem
    
    Set lstItem = lstvExplorer.ListItems.Add(, , srcRecords.Fields("FileName"))
    lstItem.Icon = ExtractIcon(srcRecords.Fields("FileName"), imgDocIco32, picTemp, 32)
    lstItem.SmallIcon = ExtractIcon(srcRecords.Fields("FileName"), imgDocIco16, picTemp, 16)

    lstItem.Tag = LAST_GENERATED_IDENTITY
    
    lstItem.SubItems(11) = FormatRecord(srcRecords.Fields("DateCreated"))
    lstItem.SubItems(12) = ""
    lstItem.SubItems(13) = srcRecords.Fields("CreatedBy")
    lstItem.SubItems(14) = ""
    lstItem.SubItems(15) = FormatRecord(srcRecords.Fields("IsConfidential"), , "IsConfidential")
    
    lstItem.ToolTipText = srcRecords.Fields("Description")
    lstItem.Selected = True
    lstItem.EnsureVisible
    
    Set lstItem = Nothing
    
    RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
    m_HaveChanges = True
    lstvExplorer.SetFocus
    
End Sub



Public Sub UpdateList(ByRef srcRecords As recordset)
    
    Dim lstItem As ListItem
    Dim iExtFld As Integer
    
    If tabDocViewer.SelectedItem.Index = 1 Then
        Set lstItem = lstvExplorer.SelectedItem
    Else
        Set lstItem = lstvSearchResult.SelectedItem
        iExtFld = 1
    End If
    
    lstItem.Icon = ExtractIcon(srcRecords.Fields("FileName"), imgDocIco32, picTemp, 32)
    lstItem.SmallIcon = ExtractIcon(srcRecords.Fields("FileName"), imgDocIco16, picTemp, 16)
    lstItem.Text = srcRecords.Fields("FileName")
    lstItem.Tag = srcRecords.Fields("ID")

    If IsNull(srcRecords.Fields("FileName")) = False Then lstItem.Text = srcRecords.Fields("FileName")
    If IsNull(srcRecords.Fields("DocumentDate")) = False Then lstItem.SubItems(5 + iExtFld) = FormatRecord(srcRecords.Fields("DocumentDate"))
    If IsNull(srcRecords.Fields("DocumentIndex")) = False Then lstItem.SubItems(4 + iExtFld) = srcRecords.Fields("DocumentIndex")
    If IsNull(srcRecords.Fields("DocumentNo")) = False Then lstItem.SubItems(3 + iExtFld) = srcRecords.Fields("DocumentNo")
    If IsNull(srcRecords.Fields("DocumentType")) = False Then lstItem.SubItems(6 + iExtFld) = srcRecords.Fields("DocumentType")
    If IsNull(srcRecords.Fields("FileSize")) = False Then lstItem.SubItems(8 + iExtFld) = FormatRecord(srcRecords.Fields("FileSize"))
    If IsNull(srcRecords.Fields("AlertDate")) = False Then lstItem.SubItems(9 + iExtFld) = FormatRecord(srcRecords.Fields("AlertDate"))
    If IsNull(srcRecords.Fields("AlertNote")) = False Then lstItem.SubItems(10 + iExtFld) = srcRecords.Fields("AlertNote")
    If IsNull(srcRecords.Fields("OriginalAuthor")) = False Then lstItem.SubItems(2 + iExtFld) = srcRecords.Fields("OriginalAuthor")
    If IsNull(srcRecords.Fields("PhysicalLocation")) = False Then lstItem.SubItems(7 + iExtFld) = srcRecords.Fields("PhysicalLocation")
    If IsNull(srcRecords.Fields("Title")) = False Then lstItem.SubItems(1) = srcRecords.Fields("Title")
    
    If IsNull(srcRecords.Fields("DateCreated")) = False Then lstItem.SubItems(11 + iExtFld) = FormatRecord(srcRecords.Fields("DateCreated"))
    If IsNull(srcRecords.Fields("LastModified")) = False Then lstItem.SubItems(12 + iExtFld) = FormatRecord(srcRecords.Fields("LastModified"))
    If IsNull(srcRecords.Fields("CreatedBy")) = False Then lstItem.SubItems(13 + iExtFld) = srcRecords.Fields("CreatedBy")
    If IsNull(srcRecords.Fields("LastModifiedBy")) = False Then lstItem.SubItems(14 + iExtFld) = srcRecords.Fields("LastModifiedBy")
    If IsNull(srcRecords.Fields("IsConfidential")) = False Then lstItem.SubItems(15 + iExtFld) = FormatRecord(srcRecords.Fields("IsConfidential"), , "IsConfidential")
    
    If IsNull(srcRecords.Fields("Description")) = False Then lstItem.ToolTipText = srcRecords.Fields("Description")
    
    lstItem.Selected = True
    lstItem.EnsureVisible
    
    Set lstItem = Nothing
    On Error Resume Next
    RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
    m_HaveChanges = True
    If tabDocViewer.SelectedItem.Index = 1 Then
        lstvExplorer.SetFocus
    Else
        lstvSearchResult.SetFocus
    End If
    
    m_HaveChanges = True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get HaveChanges() As Boolean
    HaveChanges = m_HaveChanges
End Property

Public Property Let HaveChanges(ByVal New_HaveChanges As Boolean)
    m_HaveChanges = New_HaveChanges
    PropertyChanged "HaveChanges"
End Property

Public Property Get HaveConfidentialAccess() As Boolean
    HaveConfidentialAccess = m_HaveConfidentialAccess
End Property

Public Property Let HaveConfidentialAccess(ByVal New_HaveConfidentialAccess As Boolean)
    m_HaveConfidentialAccess = New_HaveConfidentialAccess
    PropertyChanged "HaveConfidentialAccess"
End Property


Private Sub SelectAll()
    DoEvents
    
    Dim item As ListItem
    Dim recCount As Long
    
    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        For Each item In lstvExplorer.ListItems
            item.Selected = True
        Next
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        For Each item In lstvSearchResult.ListItems
            item.Selected = True
        Next
    End If

    

End Sub


Private Sub SelectNone()
    DoEvents
    
    Dim item As ListItem
    Dim recCount As Long
    
    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        For Each item In lstvExplorer.ListItems
            item.Selected = False
        Next
        
        lstvExplorer.SelectedItem.Selected = False
        Set lstvExplorer.SelectedItem = Nothing
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
        For Each item In lstvSearchResult.ListItems
            item.Selected = False
        Next
    End If

    

End Sub

Private Function GetExplorerSelItemCount() As Long
    Dim lCount As Long
    On Error GoTo err
    DoEvents
    
    Dim item As ListItem
    Dim recCount As Long
    
    If tabDocViewer.SelectedItem.Index = 1 Then
    
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Function
        
        For Each item In lstvExplorer.ListItems
            If item.Selected = True Then
                lCount = lCount + 1
            End If
        Next
        
    End If
    
    GetExplorerSelItemCount = lCount
    
    Exit Function
err:
    GetExplorerSelItemCount = 0
End Function

Private Function HaveRights() As Boolean

    Dim retVal As Boolean
    Dim recCount As Long
    
    retVal = True
    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then HaveRights = True: Exit Function
        
        If lstvExplorer.SelectedItem.SubItems(15) = "Yes" Then
            If m_HaveConfidentialAccess = False Then PrompAccessDenied: retVal = False
        End If
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then HaveRights = True: Exit Function
        
        If lstvSearchResult.SelectedItem.SubItems(16) = "Yes" Then
            If m_HaveConfidentialAccess = False Then PrompAccessDenied: retVal = False
        End If
    End If
    
    HaveRights = retVal

End Function

Private Sub MarkAsConfidentialAccess(ByVal bMark As Boolean)
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyEditFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    Dim recCount As Long
    Dim strMsg As String
    

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
        
    End If
    
    If bMark = True Then
        strMsg = "Are you sure you want to mark the selected file(s) as confidential?"
    Else
        strMsg = "Are you sure you want to mark the selected file(s) as non-confidential?"
    End If
    
    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
        
        Dim item As ListItem
    
        If tabDocViewer.SelectedItem.Index = 1 Then

            For Each item In lstvExplorer.ListItems
                If item.Selected = True Then
                    If MarkDoc(Val(item.Tag), bMark) = False Then
                        Exit For
                    Else
                        If bMark Then
                            item.SubItems(15) = "Yes"
                        Else
                            item.SubItems(15) = "No"
                        End If
                    End If
                End If
            Next
            
        Else
        
                For Each item In lstvSearchResult.ListItems
                If item.Selected = True Then
                    If MarkDoc(Val(item.Tag), bMark) = False Then
                        Exit For
                    Else
                        If bMark Then
                            item.SubItems(16) = "Yes"
                        Else
                            item.SubItems(16) = "No"
                        End If
                    End If
                End If
            Next
        
        End If
        
        m_HaveChanges = True
        HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    End If


End Sub

Private Function MarkDoc(ByVal strDocId As Long, ByVal bMark As Boolean) As Boolean
       
    Dim rsMark As recordset
    
    Set rsMark = GetRecords("SELECT [IsConfidential] FROM tbl_Files WHERE [ID]=" & strDocId)
    
    rsMark.Fields("IsConfidential") = bMark
    
    If SaveRecord("", rsMark, , True) = 1 Then
        MarkDoc = True
    End If
    Set rsMark = Nothing
    
    'MarkDoc = ExecSQL("UPDATE tbl_Files SET [IsConfidential]=" & CLng(bMark) & " WHERE [ID]=" & strDocId & "")

End Function


Private Sub ExportDoc()
    If AppCurrentUser.bCanExport = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyFileExport = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    If HaveFolderLevelAccess(GetSrcFileID(), "export") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    Dim recCount As Long
    Dim strDestPath As String
    Dim strFirstItemName As String

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
    End If
    
    strDestPath = BrowserFolder(frmFileManager.hwnd, "Please select a folder where you want to export the file(s).")
    If strDestPath = "" Then Exit Sub
    
    Dim item As ListItem
    Dim iExportCount As Integer
    
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    

    If tabDocViewer.SelectedItem.Index = 1 Then
        For Each item In lstvExplorer.ListItems
            If item.Selected = True Then
                If item.SubItems(15) = "Yes" Then
                    If item.SubItems(18) = "" Then
                        If m_HaveConfidentialAccess = True Then
                            If ExportFile(Val(item.Tag), strDestPath) = False Then
                                Exit For
                            Else
                                iExportCount = iExportCount + 1
                                If strFirstItemName = "" Then strFirstItemName = item.Text
                            End If
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                Else
                    If item.SubItems(18) = "" Then
                        If ExportFile(Val(item.Tag), strDestPath) = False Then
                            Exit For
                        Else
                            iExportCount = iExportCount + 1
                            If strFirstItemName = "" Then strFirstItemName = item.Text
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            End If
        Next
    Else
        For Each item In lstvSearchResult.ListItems
            If item.Selected = True Then
                If item.SubItems(16) = "Yes" Then
                    If item.SubItems(19) = "" Then
                        If m_HaveConfidentialAccess = True Then
                            If ExportFile(Val(item.Tag), strDestPath) = False Then
                                Exit For
                            Else
                                iExportCount = iExportCount + 1
                                If strFirstItemName = "" Then strFirstItemName = item.Text
                            End If
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                Else
                    If item.SubItems(19) = "" Then
                        If ExportFile(Val(item.Tag), strDestPath) = False Then
                            Exit For
                        Else
                            iExportCount = iExportCount + 1
                            If strFirstItemName = "" Then strFirstItemName = item.Text
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            End If
        Next
        
    End If
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    'frmWaiting.WaitingOnTop False
    
    If iExportCount > 0 Then
        MsgBox "You successfully exported " & iExportCount & " file(s).", vbInformation
        If strFirstItemName = "" Then
            Shell "Explorer.exe " & strDestPath, vbMaximizedFocus
        Else
            Shell "Explorer.exe /Select, " & strDestPath & "\" & strFirstItemName, vbMaximizedFocus
        End If
    End If

End Sub

Private Function ExportFile(ByVal strDocId As Long, ByVal strDestPath As String) As Boolean
    Dim strDocName As String
    Dim strDocFileLocation As String
    Dim rsDocument As recordset
    Dim retVal As Boolean
    
    Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & strDocId & "")
            
    If Not rsDocument Is Nothing Then
        If rsDocument.RecordCount > 0 Then
        
            strDocName = rsDocument.Fields("FileName")
            strDocFileLocation = strDestPath & "\" & strDocName
            
            Do While FileExists(strDocFileLocation) = True
                strDocFileLocation = "_" & strDocFileLocation
            Loop
        
            retVal = DownloadFileFromDB(strDocFileLocation, rsDocument, "FileObj") = True
        End If
    End If
    
    Set rsDocument = Nothing
    strDocName = vbNullString
    strDocFileLocation = vbNullString
    
    ExportFile = retVal
End Function

Public Sub ViewExpiry(ByVal startFrom As Date, ByVal endTo As Date)
    tabDocViewer.SelectedItem = tabDocViewer.Tabs(2)
    cmbFields.ListIndex = 10
    cbxCriteria.ListIndex = 7
    dtpDate(0).Value = startFrom
    dtpDate(1).Value = endTo
    cmdSearch_Click
End Sub


Private Sub UpdateBatch()
    If AppCurrentUser.bCanEdit = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyEditFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    
    Dim recCount As Long

    recCount = lstvExplorer.ListItems.Count
    If recCount = 0 Then Beep: Exit Sub
    
    Set LastRecordsetA = Nothing
    frmBatchUpdate.Show vbModal, frmFileManager

    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
    
    If Not LastRecordsetA Is Nothing Then
        Dim item As ListItem
        
        For Each item In lstvExplorer.ListItems
            If item.Selected = True Then
                If item.SubItems(15) = "Yes" Then
                    If m_HaveConfidentialAccess = True Then GoTo UPDATE_REC
                Else
UPDATE_REC:
                    If UpdateBatchOneByOne(Val(item.Tag)) = False Then
                        Exit For
                    Else
                        If item.SubItems(18) = "" Then
                            
                            On Error Resume Next
                            'Set the default
                            item.ToolTipText = ""
                            item.SubItems(5) = ""
                            item.SubItems(4) = ""
                            item.SubItems(3) = ""
                            item.SubItems(6) = ""
                            item.SubItems(9) = ""
                            item.SubItems(10) = ""
                            item.SubItems(2) = ""
                            item.SubItems(7) = ""
                            item.SubItems(7) = ""
                            item.SubItems(1) = ""
                            
                            item.SubItems(15) = ""
                            
                            item.SubItems(14) = ""
                            item.SubItems(12) = ""
                         
                         
                            If IsNull(LastRecordsetA("Description")) = False Then item.ToolTipText = LastRecordsetA("Description")
                            If IsNull(LastRecordsetA("DocumentDate")) = False Then item.SubItems(5) = LastRecordsetA("DocumentDate")
                            If IsNull(LastRecordsetA("DocumentIndex")) = False Then item.SubItems(4) = LastRecordsetA("DocumentIndex")
                            If IsNull(LastRecordsetA("DocumentNo")) = False Then item.SubItems(3) = LastRecordsetA("DocumentNo")
                            If IsNull(LastRecordsetA("DocumentType")) = False Then item.SubItems(6) = LastRecordsetA("DocumentType")
                            If IsNull(LastRecordsetA("Expiry")) = False Then item.SubItems(9) = LastRecordsetA("Expiry")
                            If IsNull(LastRecordsetA("ExpiryNote")) = False Then item.SubItems(10) = LastRecordsetA("ExpiryNote")
                            If IsNull(LastRecordsetA("OriginalAuthor")) = False Then item.SubItems(2) = LastRecordsetA("OriginalAuthor")
                            If IsNull(LastRecordsetA("PhysicalLocation")) = False Then item.SubItems(7) = LastRecordsetA("PhysicalLocation")
                            If IsNull(LastRecordsetA("Status")) = False Then item.SubItems(7) = LastRecordsetA("Status")
                            If IsNull(LastRecordsetA("Title")) = False Then item.SubItems(1) = LastRecordsetA("Title")
                            
                            If CBool(LastRecordsetA("IsConfidential")) Then
                                item.SubItems(15) = "Yes"
                            Else
                                item.SubItems(15) = "No"
                            End If
                            
                            If IsNull(LastRecordsetA("LastModifiedBy")) = False Then item.SubItems(14) = LastRecordsetA("LastModifiedBy")
                            If IsNull(LastRecordsetA("LastModified")) = False Then item.SubItems(12) = LastRecordsetA("LastModified")
                         
                        End If
                    End If
                End If
            End If
        Next
        
        m_HaveChanges = True
    
    End If
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    Set LastRecordsetA = Nothing


End Sub

Private Function UpdateBatchOneByOne(ByVal DocId As Long) As Boolean
       
    Dim strSQLUpdate As String
    
    strSQLUpdate = "SELECT [Title],[Description],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[AlertDate],[AlertNote],[IsConfidential],[Status],[LastModified],[LastModifiedBy] FROM tbl_Files WHERE (([ID]=" & DocId & ") AND (([DateCheckOut]='') OR ([DateCheckOut] IS NULL)) )"
    'strSQLUpdate = Replace(LastRecordsetA.Source, "WHERE 1=0", "WHERE (([ID]=" & DocId & ") AND (([DateCheckOut]='') OR ([DateCheckOut] IS NULL)) )")
    
    If SaveRecord(strSQLUpdate, LastRecordsetA, , True, , True) = 1 Then
        UpdateBatchOneByOne = True
    Else
    End If
    
End Function


Private Sub DuplicateDoc()
    If AppCurrentUser.bCanAdd = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyCreateFile = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    Dim recCount As Long
    

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
    End If
    
    Dim item As ListItem
    Dim i As Integer
    Dim lItemCount As Long
    

    lItemCount = lstvExplorer.ListItems.Count
    
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
            
    For i = 1 To lItemCount
    
        Set item = lstvExplorer.ListItems(i)
        If item.Selected = True Then
            If item.SubItems(15) = "Yes" Then
                If m_HaveConfidentialAccess = True Then
                    If DuplicateDocOneByOne(Val(item.Tag)) = False Then
                        Exit For
                    Else
                        item.Selected = False
                    End If
                End If
            Else
                If DuplicateDocOneByOne(Val(item.Tag)) = False Then
                    Exit For
                Else
                    item.Selected = False
                End If
            End If
        End If
        
    Next i
    
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    
    m_HaveChanges = True
       
End Sub

Private Function DuplicateDocOneByOne(ByVal DocId As String) As Boolean

    On Error GoTo err

    Dim rsDup As recordset
    Dim strFileName As String
    
    Set rsDup = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & DocId)

    If Not rsDup Is Nothing Then
        
        strFileName = "Copy of " & rsDup.Fields("FileName")
        
        Do While GetRecordCount("SELECT [ID] FROM tbl_Files WHERE (([FolderID]=" & rsDup.Fields("FolderID") & ") AND ([FileName]='" & strFileName & "'))") > 0
            strFileName = "Copy of " & strFileName
        Loop
        
        With rsDup
                    
            .Fields("FileName") = strFileName

            .Fields("LastModified") = Null
            .Fields("LastModifiedBy") = ""
            .Fields("DateCreated") = Now
            .Fields("CreatedBy") = AppCurrentUser.CompleteName
            
            '.Update
            
        End With
        
        If SaveRecord("SELECT [FileName],[Description],[FileSize],[FileObj],[FolderID],[DateCreated],[CreatedBy] FROM tbl_Files " & _
                      "WHERE [ID]=-1", rsDup, , , "tbl_Files") = 1 Then
            
            On Error Resume Next
            Dim lstItem As ListItem
            
            Set lstItem = lstvExplorer.ListItems.Add(, , rsDup.Fields("FileName"))
            lstItem.Icon = ExtractIcon(rsDup.Fields("FileName"), imgDocIco32, picTemp, 32)
            lstItem.SmallIcon = ExtractIcon(rsDup.Fields("FileName"), imgDocIco16, picTemp, 16)
            lstItem.Text = rsDup.Fields("FileName")
            lstItem.Tag = rsDup.Fields("ID")
        
            lstItem.Text = rsDup.Fields("FileName")
            lstItem.SubItems(5) = FormatRecord(rsDup.Fields("DocumentDate"))
            lstItem.SubItems(4) = rsDup.Fields("DocumentIndex")
            lstItem.SubItems(3) = rsDup.Fields("DocumentNo")
            lstItem.SubItems(6) = rsDup.Fields("DocumentType")
            lstItem.SubItems(8) = FormatRecord(rsDup.Fields("FileSize"))
            lstItem.SubItems(9) = FormatRecord(rsDup.Fields("AlertDate"))
            lstItem.SubItems(10) = rsDup.Fields("AlertNote")
            lstItem.SubItems(2) = rsDup.Fields("OriginalAuthor")
            lstItem.SubItems(7) = rsDup.Fields("PhysicalLocation")
            lstItem.SubItems(1) = rsDup.Fields("Title")

            lstItem.SubItems(11) = FormatRecord(rsDup.Fields("DateCreated"))
            lstItem.SubItems(12) = FormatRecord(rsDup.Fields("LastModified"))
            lstItem.SubItems(13) = rsDup.Fields("CreatedBy")
            lstItem.SubItems(14) = rsDup.Fields("LastModifiedBy")
            lstItem.SubItems(15) = FormatRecord(rsDup.Fields("IsConfidential"), , "IsConfidential")
            
            lstItem.ToolTipText = rsDup.Fields("Description")
            
            lstItem.Selected = True
            lstItem.EnsureVisible
            
            'Set lstItem = Nothing
            
            RaiseEvent NavChange("Record: " & lstvExplorer.SelectedItem.Index & " of " & lstvExplorer.ListItems.Count)
            m_HaveChanges = True
            lstvExplorer.SetFocus
            
        End If
    
        Set rsDup = Nothing
        
        DuplicateDocOneByOne = True
    End If
        
err:
    
End Function

Public Sub CheckIn()
    If AppCurrentUser.bCanChkOut = False Then PrompAccessDenied: Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyCheckOut = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    If HaveFolderLevelAccess(GetSrcFileID(), "checkin") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
On Error GoTo cError
    
    Dim recCount As Long
    Dim strDocId As Long
    Dim strDocLocation As String
    
    recCount = lstvExplorer.ListItems.Count
    If recCount = 0 Then Beep: Exit Sub
    If lstvExplorer.SelectedItem.SubItems(19) = "" Then Beep: Exit Sub
    
    If AppCurrentUser.bIsSysAdmin = False Then
        If Val(lstvExplorer.SelectedItem.SubItems(19)) <> AppCurrentUser.UserId Then
            MsgBox "Only '" & lstvExplorer.SelectedItem.SubItems(17) & "' or The Administrator can check-in the file.", vbInformation
            Exit Sub
        End If
    End If
    
    With dlgCheckIn
        .MaxFileSize = 32000 'this will max out the buffer for the filenames array for large selections. *NEW*
        .DialogTitle = "File Check-in"
        .fileName = ""
        .Filter = "*." & GetFileExt(lstvExplorer.SelectedItem.Text) & "|*." & GetFileExt(lstvExplorer.SelectedItem.Text)
        .ShowOpen
        
        If .fileName = "" Then
            Exit Sub
        Else
            strDocLocation = .fileName
            
            strDocId = Val(lstvExplorer.SelectedItem.Tag)

            If MsgBox("Are you sure you want to checkin the file?", vbQuestion + vbYesNo, "Confirm File Check-in") = vbYes Then
                DoEvents
                ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
                

                If PerformCheckIn(strDocId, strDocLocation) = True Then
                    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
                    MsgBox "File has been successfully checked-in.", vbInformation
                Else
                    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
                End If
                
                
            End If
            
        End If

    End With

    Exit Sub
cError:
    MsgBox err.Description, vbCritical, "Unexpected Error"
    
    
End Sub

Private Function PerformCheckIn(ByVal strDocId As String, ByVal strDocLocation As String, Optional removeUnwantedFileString As String, Optional importDesc As String) As Boolean
    
    Dim retVal As Boolean
    
    
    
    If strDocLocation <> "" Then
    
        Dim rsFiles As recordset
        Dim strDocFileName As String
        
        Set rsFiles = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & strDocId)
        
        strDocFileName = GetFileNameFromPath(strDocLocation)
        strDocFileName = Replace(strDocFileName, "_" & removeUnwantedFileString, "")
    
        If Not rsFiles Is Nothing Then
            With rsFiles
                
                .Fields("FileName") = strDocFileName
                
                SaveFileToDB strDocLocation, rsFiles, "FileObj"

                .Fields("FileSize") = GetFileSizeString(strDocLocation)
                .Fields("FileSizeNo") = GetFileSize(strDocLocation)
                .Fields("LastModified") = Now
                .Fields("LastModifiedBy") = AppCurrentUser.CompleteName
                
                '.Update
                
            End With
            
            If SaveRecord("SELECT [FileName],[FileSize],[FileSizeNo],[FileObj],[LastModified],[LastModifiedBy] FROM tbl_Files " & _
                          "WHERE [ID]=" & strDocId, rsFiles, , True, "tbl_Files") = 1 Then

                Dim rsChkIn As recordset
                Set rsChkIn = GetRecords("SELECT [CheckOutByFK],[CheckOutBy],[DateCheckOut] FROM tbl_Files WHERE (([ID]=" & strDocId & "))") ' AND ([CheckOutByFK]='" & AppCurrentUser.UserId & "'))")
                
                rsChkIn.Fields("CheckOutByFK") = Null
                rsChkIn.Fields("CheckOutBy") = ""
                rsChkIn.Fields("DateCheckOut") = ""
                
                If SaveRecord("", rsChkIn, , True) = 1 Then
                    retVal = True
                End If
                Set rsChkIn = Nothing
                
                'retVal = ExecSQL("UPDATE tbl_Files SET [CheckOutByFK]=NULL,[CheckOutBy]='',[DateCheckOut]='' WHERE (([ID]=" & strDocId & "))") ' AND ([CheckOutByFK]='" & AppCurrentUser.UserId & "'))")
                
                If retVal = True Then
                    lstvExplorer.SelectedItem.Text = strDocFileName
                    
                    lstvExplorer.SelectedItem.SubItems(8) = rsFiles.Fields("FileSize")
                    lstvExplorer.SelectedItem.SubItems(12) = rsFiles.Fields("LastModified")
                    lstvExplorer.SelectedItem.SubItems(14) = rsFiles.Fields("LastModifiedBy")
                    
                    lstvExplorer.SelectedItem.SubItems(17) = ""
                    lstvExplorer.SelectedItem.SubItems(18) = ""
                    lstvExplorer.SelectedItem.SubItems(19) = ""
                End If
                
            End If
        
            Set rsFiles = Nothing
        End If
    End If
            
    PerformCheckIn = retVal
    
End Function

Public Sub CheckOut()
    If AppCurrentUser.bCanChkOut = False Then PrompAccessDenied: Exit Sub
    If HaveRights = False Then Exit Sub
    If CurrentFolderAccess.folderId = m_FolderId And CurrentFolderAccess.bDenyCheckOut = True Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    If HaveFolderLevelAccess(GetSrcFileID(), "checkout") = False Then
        PrompAccessDeniedForFolder
        Exit Sub
    End If
    
    DoEvents
    
    Dim recCount As Long
    Dim strDestPath As String
    Dim strFirstItemName As String

    If tabDocViewer.SelectedItem.Index = 1 Then
        recCount = lstvExplorer.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
    Else
        recCount = lstvSearchResult.ListItems.Count
        If recCount = 0 Then Beep: Exit Sub
    End If
    
    strDestPath = BrowserFolder(frmFileManager.hwnd, "Please select a folder where you want to check out the file(s).")
    If strDestPath = "" Then Exit Sub
    
    Dim item As ListItem
    Dim iCheckOutCount As Integer
    
    ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    

    If tabDocViewer.SelectedItem.Index = 1 Then
        For Each item In lstvExplorer.ListItems
            If item.Selected = True Then
                If item.SubItems(15) = "Yes" Then
                    If m_HaveConfidentialAccess = True Then
                        If item.SubItems(18) = "" Then
                            If CheckOutFiles(Val(item.Tag), strDestPath, item) = False Then
                                Exit For
                            Else
                                iCheckOutCount = iCheckOutCount + 1
                                If strFirstItemName = "" Then strFirstItemName = item.Text
                            End If
                         Else
                            'frmWaiting.WaitingOnTop False
                            MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                            'frmWaiting.WaitingOnTop False
                        End If
                    End If
                Else
                    If item.SubItems(18) = "" Then
                        If CheckOutFiles(Val(item.Tag), strDestPath, item) = False Then
                            Exit For
                        Else
                            iCheckOutCount = iCheckOutCount + 1
                            If strFirstItemName = "" Then strFirstItemName = item.Text
                        End If
                     Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            End If
        Next
    Else
        For Each item In lstvSearchResult.ListItems
            If item.Selected = True Then
                If item.SubItems(16) = "Yes" Then
                    If m_HaveConfidentialAccess = True Then
                        If item.SubItems(19) = "" Then
                            If CheckOutFiles(Val(item.Tag), strDestPath, item) = False Then
                                Exit For
                            Else
                                iCheckOutCount = iCheckOutCount + 1
                                If strFirstItemName = "" Then strFirstItemName = item.Text
                            End If
                         Else
                            'frmWaiting.WaitingOnTop False
                            MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                            'frmWaiting.WaitingOnTop False
                        End If
                    End If
                Else
                    If item.SubItems(19) = "" Then
                        If CheckOutFiles(Val(item.Tag), strDestPath, item) = False Then
                            Exit For
                        Else
                            iCheckOutCount = iCheckOutCount + 1
                            If strFirstItemName = "" Then strFirstItemName = item.Text
                        End If
                    Else
                        'frmWaiting.WaitingOnTop False
                        MsgBox "The file named '" & item.Text & "' is currently check-out.", vbInformation
                        'frmWaiting.WaitingOnTop False
                    End If
                End If
            End If
        Next
        
    End If
    
        
    HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
    'frmWaiting.WaitingOnTop False
    
    If iCheckOutCount > 0 Then
        MsgBox "You successfully check-out " & iCheckOutCount & " file(s).", vbInformation
        
        If strFirstItemName = "" Then
            Shell "Explorer.exe " & strDestPath, vbMaximizedFocus
        Else
            Shell "Explorer.exe /Select, " & strDestPath & "\" & strFirstItemName, vbMaximizedFocus
        End If
    End If

End Sub


Private Function CheckOutFiles(ByVal strDocId As Long, ByVal strDestPath As String, ByRef item As ListItem) As Boolean
    Dim strDocName As String
    Dim strDocFileLocation As String
    Dim rsDocument As recordset
    Dim retVal As Boolean
    
    Set rsDocument = GetRecords("SELECT * FROM tbl_Files WHERE [ID]=" & strDocId & "")
            
    If Not rsDocument Is Nothing Then
        If rsDocument.RecordCount > 0 Then
        
            strDocName = rsDocument.Fields("FileName")
            strDocFileLocation = strDestPath & "\" & strDocName
            
            Do While FileExists(strDocFileLocation) = True
                strDocFileLocation = "_" & strDocFileLocation
            Loop
        
            retVal = DownloadFileFromDB(strDocFileLocation, rsDocument, "FileObj") = True
            
            Dim rsChkOut As recordset
            If retVal = True Then
                'Update for check out
                If tabDocViewer.SelectedItem.Index = 1 Then
                    item.SubItems(17) = AppCurrentUser.CompleteName
                    item.SubItems(18) = Now
                    item.SubItems(19) = strDocId
                    
                    
                    Set rsChkOut = GetRecords("SELECT [CheckOutByFK],[CheckOutBy],[DateCheckOut] FROM tbl_Files WHERE (([ID]=" & item.SubItems(19) & "))") ' AND ([CheckOutByFK]='" & AppCurrentUser.UserId & "'))")
                    
                    rsChkOut.Fields("CheckOutByFK") = AppCurrentUser.UserId
                    rsChkOut.Fields("CheckOutBy") = item.SubItems(17)
                    rsChkOut.Fields("DateCheckOut") = item.SubItems(18)
                    
                    If SaveRecord("", rsChkOut, , True) = 1 Then
                        retVal = True
                    End If
                    Set rsChkOut = Nothing
                
                    'retVal = ExecSQL("UPDATE tbl_Files SET [CheckOutByFK]=" & AppCurrentUser.UserId & ",[CheckOutBy]='" & item.SubItems(17) & "',[DateCheckOut]='" & item.SubItems(18) & "' WHERE [ID]=" & item.SubItems(19) & "")
                Else
                    item.SubItems(18) = AppCurrentUser.CompleteName
                    item.SubItems(19) = Now
                    item.SubItems(20) = strDocId
                    
                    Set rsChkOut = GetRecords("SELECT [CheckOutByFK],[CheckOutBy],[DateCheckOut] FROM tbl_Files WHERE (([ID]=" & item.SubItems(20) & "))") ' AND ([CheckOutByFK]='" & AppCurrentUser.UserId & "'))")
                    
                    rsChkOut.Fields("CheckOutByFK") = AppCurrentUser.UserId
                    rsChkOut.Fields("CheckOutBy") = item.SubItems(18)
                    rsChkOut.Fields("DateCheckOut") = item.SubItems(19)
                    
                    If SaveRecord("", rsChkOut, , True) = 1 Then
                        retVal = True
                    End If
                    Set rsChkOut = Nothing
                    
                    'retVal = ExecSQL("UPDATE tbl_Files SET [CheckOutByFK]=" & AppCurrentUser.UserId & ",[CheckOutBy]='" & item.SubItems(18) & "',[DateCheckOut]='" & item.SubItems(19) & "' WHERE [ID]=" & item.SubItems(20) & "")
                End If
            End If
            
        End If
    End If
    
    Set rsDocument = Nothing
    strDocName = vbNullString
    strDocFileLocation = vbNullString
    
    CheckOutFiles = retVal
End Function


Public Sub Reset()
    m_FolderId = 0
    
    lstvExplorer.ListItems.Clear
    lstvSearchResult.ListItems.Clear
    
    RaiseEvent NavChange("Record: 0 of 0")
End Sub

Private Function GetSrcFileID() As Long
    On Error Resume Next
    GetSrcFileID = Val(lstvSearchResult.SelectedItem.Tag)
End Function

Private Function HaveFolderLevelAccess(ByVal FileID As Long, ByVal AccessName As String) As Boolean
    
    Dim retVal As Boolean
    retVal = True
    
    If tabDocViewer.SelectedItem.Index <> 1 Then
        If FileID <> 0 Then
            If AppCurrentUser.bIsSysAdmin = False Then
                ShowWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
                
        
                Dim rsRecord As recordset
                Dim lFolderId As Long
                
                lFolderId = Val(GetValueAt("SELECT [FolderID] FROM tbl_Files WHERE [ID]=" & FileID, "FolderID"))
    
                If lFolderId > 0 Then
                    Set rsRecord = GetRecords("SELECT * FROM vw_FolderRestrictionSet WHERE [FolderID]=" & lFolderId & " AND [UserGroupID]=" & AppCurrentUser.UserGroupId)
                    If Not rsRecord Is Nothing Then
                        If rsRecord.RecordCount > 0 Then
                        
                            Select Case AccessName
                                Case "open"
                                    If Val(rsRecord.Fields("DenyOpenFile")) <> 0 Then retVal = False
                                Case "edit"
                                    If Val(rsRecord.Fields("DenyEditFile")) <> 0 Then retVal = False
                                Case "delete"
                                    If Val(rsRecord.Fields("DenyDeleteFile")) <> 0 Then retVal = False
                                Case "export"
                                    If Val(rsRecord.Fields("DenyFileExport")) <> 0 Then retVal = False
                                Case "checkout"
                                    If Val(rsRecord.Fields("DenyCheckOut")) <> 0 Then retVal = False
                            End Select
                            
                        End If
                    End If
                End If
                Set rsRecord = Nothing
                
                HideWaiting UserControl.Parent, UserControl.Parent.WaitingDisplay
            End If
        End If
    End If

    HaveFolderLevelAccess = retVal

End Function
