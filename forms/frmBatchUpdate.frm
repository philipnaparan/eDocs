VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatchUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Update"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatchUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbxStatus 
      DataField       =   "Status"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1650
      TabIndex        =   23
      Top             =   6525
      Width           =   3540
   End
   Begin VB.CheckBox Check12 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc. Status"
      Height          =   240
      Left            =   150
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6525
      Width           =   1440
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -6675
      TabIndex        =   28
      Top             =   7050
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   53
   End
   Begin VB.TextBox txtPhysicalLoc 
      DataField       =   "PhysicalLocation"
      Height          =   315
      Left            =   1650
      TabIndex        =   15
      Top             =   4350
      Width           =   3540
   End
   Begin VB.CheckBox Check11 
      Alignment       =   1  'Right Justify
      Caption         =   "Physical Loc."
      Height          =   240
      Left            =   150
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4350
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CheckBox Check10 
      Alignment       =   1  'Right Justify
      Caption         =   "Title"
      Height          =   240
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1440
   End
   Begin VB.TextBox txtTitle 
      DataField       =   "Title"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1650
      TabIndex        =   1
      Top             =   1050
      Width           =   3540
   End
   Begin VB.CheckBox Check9 
      Alignment       =   1  'Right Justify
      Caption         =   "Marking"
      Height          =   240
      Left            =   1650
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6225
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.CheckBox Check8 
      Alignment       =   1  'Right Justify
      Caption         =   "Alert Note"
      Height          =   240
      Left            =   150
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1440
   End
   Begin VB.CheckBox Check7 
      Alignment       =   1  'Right Justify
      Caption         =   "Alert Date"
      Height          =   240
      Left            =   150
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1440
   End
   Begin VB.CheckBox Check6 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc. Type"
      Height          =   240
      Left            =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1440
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc. Date"
      Height          =   240
      Left            =   150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc. Index"
      Height          =   240
      Left            =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3225
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Doc. No"
      Height          =   240
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2850
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Original Author"
      Height          =   240
      Left            =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2475
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1425
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.TextBox txtEmpNote 
      DataField       =   "AlertNote"
      Enabled         =   0   'False
      Height          =   990
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   5175
      Width           =   3540
   End
   Begin VB.CheckBox ckConfidential 
      Caption         =   "Mark As Confidential"
      DataField       =   "IsConfidential"
      Height          =   285
      Left            =   3300
      TabIndex        =   21
      Top             =   6195
      Width           =   1860
   End
   Begin VB.ComboBox cbxDocType 
      DataField       =   "DocumentType"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1650
      TabIndex        =   13
      Top             =   3975
      Width           =   3540
   End
   Begin MSComCtl2.DTPicker dtpDocDate 
      DataField       =   "DocumentDate"
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   3600
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   61407233
      CurrentDate     =   39436
   End
   Begin VB.TextBox txtIndex 
      DataField       =   "DocumentIndex"
      Height          =   315
      Left            =   1650
      TabIndex        =   9
      Top             =   3225
      Width           =   3540
   End
   Begin VB.TextBox txtNo 
      DataField       =   "DocumentNo"
      Height          =   315
      Left            =   1650
      TabIndex        =   7
      Top             =   2850
      Width           =   3540
   End
   Begin VB.TextBox txtAuthor 
      DataField       =   "OriginalAuthor"
      Height          =   315
      Left            =   1650
      TabIndex        =   5
      Top             =   2475
      Width           =   3540
   End
   Begin VB.TextBox txtDesc 
      DataField       =   "Description"
      Height          =   990
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1425
      Width           =   3540
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   390
      Left            =   3975
      TabIndex        =   24
      Top             =   7200
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   2625
      TabIndex        =   26
      Top             =   7200
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpDocExp 
      DataField       =   "AlertDate"
      Height          =   315
      Left            =   1650
      TabIndex        =   17
      Top             =   4800
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      CheckBox        =   -1  'True
      Format          =   61407233
      CurrentDate     =   39436
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   29
      Top             =   600
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1275
      TabIndex        =   30
      Top             =   8400
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmBatchUpdate.frx":038A
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Check the field that you want to update:"
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
      Left            =   150
      TabIndex        =   27
      Top             =   750
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Update"
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
      TabIndex        =   25
      Top             =   150
      Width           =   2115
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
Attribute VB_Name = "frmBatchUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check1_Click()
    txtDesc.Enabled = Not txtDesc.Enabled
End Sub

Private Sub Check10_Click()
    txtTitle.Enabled = Not txtTitle.Enabled
End Sub

Private Sub Check11_Click()
    txtPhysicalLoc.Enabled = Not txtPhysicalLoc.Enabled
End Sub

Private Sub Check12_Click()
    cbxStatus.Enabled = Not cbxStatus.Enabled
End Sub

Private Sub Check2_Click()
    txtAuthor.Enabled = Not txtAuthor.Enabled
End Sub

Private Sub Check3_Click()
    txtNo.Enabled = Not txtNo.Enabled
End Sub

Private Sub Check4_Click()
    txtIndex.Enabled = Not txtIndex.Enabled
End Sub

Private Sub Check5_Click()
    dtpDocDate.Enabled = Not dtpDocDate.Enabled
End Sub

Private Sub Check6_Click()
    cbxDocType.Enabled = Not cbxDocType.Enabled
End Sub

Private Sub Check7_Click()
    dtpDocExp.Enabled = Not dtpDocExp.Enabled
End Sub

Private Sub Check8_Click()
    txtEmpNote.Enabled = Not txtEmpNote.Enabled
End Sub

Private Sub Check9_Click()
    ckConfidential.Enabled = Not ckConfidential.Enabled
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ShowWaiting Me, WaitingDisplay
    Set LastRecordsetA = GetRecords("SELECT [Title],[Description],[OriginalAuthor],[DocumentNo],[DocumentIndex],[DocumentDate],[DocumentType],[PhysicalLocation],[AlertDate],[AlertNote],[IsConfidential],[Status],[LastModified],[LastModifiedBy] FROM tbl_Files WHERE 1=0")
    HideWaiting Me, WaitingDisplay
    
    If Not LastRecordsetA Is Nothing Then
        LastRecordsetA.AddNew
        
        If txtTitle.Enabled Then LastRecordsetA.Fields("Title") = txtTitle.Text
        If txtDesc.Enabled Then LastRecordsetA.Fields("Description") = txtDesc.Text
        If txtAuthor.Enabled Then LastRecordsetA.Fields("OriginalAuthor") = txtAuthor.Text
        If txtNo.Enabled Then LastRecordsetA.Fields("DocumentNo") = txtNo.Text
        If txtIndex.Enabled Then LastRecordsetA.Fields("DocumentIndex") = txtIndex.Text
        If dtpDocDate.Enabled Then
            If IsNull(dtpDocDate.Value) Then
                LastRecordsetA.Fields("DocumentDate") = Null
            Else
                LastRecordsetA.Fields("DocumentDate") = CDate(dtpDocDate.Value)
            End If
        End If
        If cbxDocType.Enabled Then LastRecordsetA.Fields("DocumentType") = cbxDocType.Text
        If dtpDocExp.Enabled Then
            If IsNull(dtpDocExp.Value) Then
                LastRecordsetA.Fields("AlertDate") = Null
            Else
                LastRecordsetA.Fields("AlertDate") = CDate(dtpDocExp.Value)
            End If
        End If
        If txtEmpNote.Enabled Then LastRecordsetA.Fields("AlertNote") = txtEmpNote.Text
        If txtPhysicalLoc.Enabled Then LastRecordsetA.Fields("PhysicalLocation") = txtPhysicalLoc.Text
        If ckConfidential.Enabled Then LastRecordsetA.Fields("IsConfidential") = ckConfidential.Value
        If cbxStatus.Enabled Then LastRecordsetA.Fields("Status") = cbxStatus.Text
        
        LastRecordsetA.Fields("LastModified") = Now
        LastRecordsetA.Fields("LastModifiedBy") = AppCurrentUser.CompleteName
        
        
    Else
        MsgBox "Unexpected error.", vbCritical
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    ShowWaiting Me, WaitingDisplay
    
    dtpDocDate.Value = Date
    dtpDocExp.Value = Date
    
    Dim rs As recordset
    
    Set rs = GetRecords("SELECT * FROM tbl_DocType ORDER BY TypeName ASC")
    If Not rs Is Nothing Then
        FillComboBox rs, cbxDocType, "TypeName"
        Set rs = Nothing
    End If
    
    Set rs = GetRecords("SELECT * FROM tbl_FileStatus ORDER BY StatusName ASC")
    If Not rs Is Nothing Then
        FillComboBox rs, cbxStatus, "StatusName"
        Set rs = Nothing
    End If
    
    cbxDocType.AddItem "Others"
    
    HideWaiting Me, WaitingDisplay
    Screen.MousePointer = vbDefault
End Sub

