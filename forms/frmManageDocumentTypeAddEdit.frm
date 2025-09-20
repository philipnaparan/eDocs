VERSION 5.00
Begin VB.Form frmManageDocumentTypeAddEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2235
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
   Icon            =   "frmManageDocumentTypeAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2430
      TabIndex        =   1
      Top             =   1605
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3645
      TabIndex        =   2
      Top             =   1605
      Width           =   1140
   End
   Begin VB.TextBox txtName 
      DataField       =   "TypeName"
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
      TabIndex        =   5
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -225
      TabIndex        =   6
      Top             =   1425
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   300
      TabIndex        =   7
      Top             =   4200
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmManageDocumentTypeAddEdit.frx":038A
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
      TabIndex        =   4
      Top             =   150
      Width           =   3690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Type Name:"
      Height          =   315
      Left            =   135
      TabIndex        =   3
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
Attribute VB_Name = "frmManageDocumentTypeAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lPK As Long

Dim bAddState As Boolean
Dim rsRecord As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(txtName) Then Exit Sub
    
    ShowWaiting Me, WaitingDisplay
    
    SyncronizeRecordsetBinding rsRecord, Me
    SaveRecord "SELECT [TypeName]" & _
           " FROM tbl_DocType WHERE [ID]=" & lPK, rsRecord, , Not bAddState, "tbl_DocType"
    
    If bAddState = True Then
        Set LastRecordsetA = GetRecords("SELECT * FROM tbl_DocType WHERE [ID]=" & LAST_GENERATED_IDENTITY)
    Else
        Set LastRecordsetA = rsRecord
    End If
    HideWaiting Me, WaitingDisplay
    Unload Me
    
End Sub

Private Sub Form_Load()
    If lPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Set rsRecord = GetRecords("SELECT * FROM tbl_DocType WHERE [ID]=" & lPK)
    If rsRecord Is Nothing Then
        MsgBox "Unable to read document type data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        BindControlToRS Me, rsRecord
    End If
    
    If bAddState = True Then
        Me.Caption = "Add New Document Type"
        rsRecord.AddNew
    Else
        Me.Caption = "Edit Document Type"
    End If
    
    lblTitle.Caption = Me.Caption
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmManageDocumentTypeAddEdit = Nothing
End Sub
