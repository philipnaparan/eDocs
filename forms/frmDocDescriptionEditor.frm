VERSION 5.00
Begin VB.Form frmDocDescriptionEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Description"
   ClientHeight    =   3090
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   5250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDocDescriptionEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3915
      TabIndex        =   1
      Top             =   2565
      Width           =   1230
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   90
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   135
      Width           =   5055
   End
End
Attribute VB_Name = "frmDocDescriptionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public bViewerOnly As Boolean

Private Sub cmdSave_Click()
'    If bViewerOnly = True Then
'        Unload Me
'    Else
'        strLastFileDesc = txtDesc.Text
'        Unload Me
'    End If
End Sub

Private Sub Form_Load()
'    If bViewerOnly = True Then
'        cmdSave.Caption = "&Close"
'    Else
'        cmdSave.Caption = "&Save"
'    End If
End Sub
