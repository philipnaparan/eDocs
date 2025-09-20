VERSION 5.00
Begin VB.Form frmFilePropertyAddEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Property Editor"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFilePropertyAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValue 
      DataField       =   "PropertyValue"
      Height          =   315
      Left            =   1875
      TabIndex        =   1
      Top             =   1275
      Width           =   3015
   End
   Begin VB.ComboBox cbxProperty 
      DataField       =   "PropertyName"
      Height          =   315
      ItemData        =   "frmFilePropertyAddEdit.frx":038A
      Left            =   1890
      List            =   "frmFilePropertyAddEdit.frx":038C
      TabIndex        =   0
      Text            =   "cbxUserType"
      Top             =   840
      Width           =   3030
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   2400
      TabIndex        =   3
      Top             =   2025
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   390
      Left            =   3750
      TabIndex        =   2
      Top             =   2025
      Width           =   1230
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -6825
      TabIndex        =   6
      Top             =   1875
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1125
      TabIndex        =   9
      Top             =   5100
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Property Name:"
      Height          =   315
      Left            =   225
      TabIndex        =   8
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Property Value:"
      Height          =   315
      Left            =   225
      TabIndex        =   7
      Top             =   1275
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Property Editor"
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
      TabIndex        =   5
      Top             =   150
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmFilePropertyAddEdit.frx":038E
      Top             =   150
      Width           =   360
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
Attribute VB_Name = "frmFilePropertyAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lRecordPK As Long
Public lFileId As Long

Dim bAddState As Boolean
Dim rsRecord As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(cbxProperty) Then Exit Sub
    If IsControlEmpty(txtValue) Then Exit Sub
    
    If IsExistInCombo(cbxProperty.Text, cbxProperty) = False Then cbxProperty.SetFocus: Exit Sub
    
    ShowWaiting Me, WaitingDisplay
    
    rsRecord.Fields("FileId") = lFileId
    
    SyncronizeRecordsetBinding rsRecord, Me
    SaveRecord "SELECT [FileId],[PropertyName],[PropertyValue] FROM tbl_FileProperty WHERE [ID]=" & lRecordPK, rsRecord, , Not bAddState
    
    HideWaiting Me, WaitingDisplay
    
    Unload Me
End Sub

Private Sub Form_Load()
    If lRecordPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Dim rsDropList As recordset
   
    Set rsDropList = GetRecords("SELECT * FROM tbl_PropertyList ORDER BY PropertyName ASC")
    If Not rsDropList Is Nothing Then
         FillComboBox rsDropList, cbxProperty, "PropertyName"
         Set rsDropList = Nothing
    End If
    
    Set rsRecord = GetRecords("SELECT * FROM tbl_FileProperty WHERE [ID]=" & lRecordPK)
    If rsRecord Is Nothing Then
        MsgBox "Unable to read the data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        BindControlToRS Me, rsRecord
    End If
    
    If bAddState = True Then
        Me.Caption = "Add New Record"
        rsRecord.AddNew
    Else
        Me.Caption = "Edit Record"
    End If
    
    
    
    lblTitle.Caption = Me.Caption
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFilePropertyAddEdit = Nothing
End Sub

