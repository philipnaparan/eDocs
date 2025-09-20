VERSION 5.00
Begin VB.Form frmShellWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   540
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2595
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
   Icon            =   "frmShellWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   900
      Top             =   300
   End
   Begin VB.Timer tmrUnloader 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5700
      Top             =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Waiting to close the file..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   2340
   End
End
Attribute VB_Name = "frmShellWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fileName As String
Dim udtShellInfo As udtShellAndWait

Private Sub Form_Load()
    tmrDelay.Enabled = True
End Sub

Public Sub OpenFile()
    With udtShellInfo
        .bLogFile = False
        .sCommand = App.Path & "\ShellExecute.exe " & fileName
    End With
    ShellAndWait udtShellInfo
    tmrUnloader.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fileName = ""
End Sub

Private Sub tmrDelay_Timer()
    tmrDelay.Enabled = False
    OpenFile
End Sub

Private Sub tmrUnloader_Timer()
    If udtShellInfo.bShellAndWaitRunning = False Then tmrUnloader.Enabled = False: Unload Me
End Sub
