VERSION 5.00
Begin VB.Form frmNewDocType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Document"
   ClientHeight    =   3225
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   3345
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
   Icon            =   "frmNewDocType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLocker 
      Interval        =   1
      Left            =   2700
      Top             =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   690
      Left            =   450
      TabIndex        =   8
      Top             =   2175
      Width           =   840
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Powerpoint Document"
      Height          =   330
      Left            =   210
      TabIndex        =   5
      Top             =   1365
      Width           =   2010
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Excel Document"
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   1050
      Width           =   2010
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Word Document"
      Height          =   330
      Left            =   210
      TabIndex        =   3
      Top             =   735
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.TextBox txtDocName 
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   2100
      Width           =   2850
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2100
      TabIndex        =   2
      Top             =   2625
      Width           =   960
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   2625
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the name of the document:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   1785
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "What kind of document you want to create?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   2850
   End
End
Attribute VB_Name = "frmNewDocType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
'Dim udtShellInfo As udtShellAndWait
'
'Private Sub btnCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub btnOK_Click()
'    If txtDocName.Text = "" Then
'        Beep
'        txtDocName.SetFocus
'        Exit Sub
'    End If
'
'    If Option1.Value = True Then LastDocInfo.DocType = "doc"
'    If Option2.Value = True Then LastDocInfo.DocType = "xls"
'    If Option3.Value = True Then LastDocInfo.DocType = "ppt"
'
'    LastDocInfo.DocName = txtDocName.Text
'
'    Unload Me
'
'End Sub
'
'
'
'Private Sub Command1_Click()
'    frmShellWait.FileName = "C:\Documents and Settings\Philip Naparan\Desktop\code\_Image_Vie3644111212001\ReadMe.txt"
'    frmShellWait.Show vbModal
'End Sub
'
'Private Sub tmrLocker_Timer()
'    Me.Enabled = Not udtShellInfo.bShellAndWaitRunning
'End Sub
