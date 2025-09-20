VERSION 5.00
Begin VB.Form frmOpenOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Open Options"
   ClientHeight    =   1275
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   4260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenOption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   390
      Left            =   2700
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpenWith 
      Caption         =   "Open &With"
      Height          =   390
      Left            =   2700
      TabIndex        =   1
      Top             =   675
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Open wiith default editor."
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   225
      Width           =   2490
   End
   Begin VB.Label Label2 
      Caption         =   "Choose an editor for the file."
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   750
      Width           =   2490
   End
End
Attribute VB_Name = "frmOpenOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
    LastOpenOption = "OPEN_DEFAULT"
    Unload Me
End Sub

Private Sub cmdOpenWith_Click()
    LastOpenOption = "OPEN_WITH"
    Unload Me
End Sub

Private Sub Form_Load()
    LastOpenOption = "OPEN_DEFAULT"
End Sub
