VERSION 5.00
Begin VB.Form frmFileAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New File"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      DataField       =   "FolderName"
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   780
      Width           =   3015
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3570
      TabIndex        =   2
      Top             =   1605
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2355
      TabIndex        =   1
      Top             =   1605
      Width           =   1140
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   3
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -300
      TabIndex        =   4
      Top             =   1425
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   75
      TabIndex        =   7
      Top             =   3750
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "File Name:"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "New File"
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
      TabIndex        =   5
      Top             =   150
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmFileAdd.frx":038A
      Top             =   150
      Width           =   360
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
Attribute VB_Name = "frmFileAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(txtName) Then Exit Sub
    
    LastGenericText = txtName.Text
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFileAdd = Nothing
End Sub


