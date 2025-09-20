VERSION 5.00
Begin VB.Form frmFileLuncher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Editor"
   ClientHeight    =   2340
   ClientLeft      =   255
   ClientTop       =   1800
   ClientWidth     =   5385
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
   Icon            =   "frmFileLuncher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3825
      TabIndex        =   2
      Top             =   1725
      Width           =   1365
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish Editing"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3825
      TabIndex        =   1
      Top             =   825
      Width           =   1365
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start Editing"
      Default         =   -1  'True
      Height          =   390
      Left            =   3825
      TabIndex        =   0
      Top             =   300
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   150
      Picture         =   "frmFileLuncher.frx":000C
      Top             =   375
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Steps:"
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
      Left            =   975
      TabIndex        =   5
      Top             =   225
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "2. After you finish editing the file, save it and then close it and press the 'Finish Editing' button."
      Height          =   690
      Left            =   975
      TabIndex        =   4
      Top             =   1350
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "1. To start editing the file press the 'Start Editing' button."
      Height          =   615
      Left            =   975
      TabIndex        =   3
      Top             =   600
      Width           =   2715
   End
End
Attribute VB_Name = "frmFileLuncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fileName As String


Private Sub cmdCancel_Click()
    fileName = ""
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    LastUseFileNamePath = fileName
    Unload Me
End Sub

Private Sub cmdStart_Click()
    
    cmdFinish.Enabled = True
    cmdFinish.Default = True
    cmdStart.Enabled = False
    
    frmOpenOption.Show vbModal
    
    If LastOpenOption = "OPEN_DEFAULT" Then
        OpenURL fileName, Me.hwnd
    Else
        LunchFileWithDialog fileName
    End If
    
End Sub

Private Sub Form_Load()
    LastUseFileNamePath = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fileName = ""
End Sub

