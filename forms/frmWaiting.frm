VERSION 5.00
Begin VB.Form frmWaiting 
   BorderStyle     =   0  'None
   ClientHeight    =   690
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   2565
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
   Icon            =   "frmWaiting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnim 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrUnloader 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1875
      Top             =   750
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   2565
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "frmWaiting.frx":000C
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblWaitingMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   900
      TabIndex        =   0
      Top             =   150
      Width           =   1740
   End
End
Attribute VB_Name = "frmWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Integer

Private Sub tmrAnim_Timer()
    i = i + 1
    If i > 3 Then i = 0
    
    If i = 0 Then
        lblWaitingMsg.Caption = "Processing"
    Else
        lblWaitingMsg.Caption = "Processing" & FillStr(".", i)
    End If
End Sub

Private Sub Form_Load()

    Me.Top = frmFileManager.Top + ((frmFileManager.Height - Me.Height) / 2)
    Me.Left = frmFileManager.Left + ((frmFileManager.Width - Me.Width) / 2)

    Screen.MousePointer = vbHourglass

    Me.Show
    OnTop Me.hwnd, True
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Screen.MousePointer = vbDefault
    
    Set frmWaiting = Nothing
End Sub

Public Sub ForceTerminate()

    tmrUnloader.Enabled = False
    Unload Me

End Sub

Public Sub Terminate()

    tmrUnloader.Enabled = True

End Sub

Private Sub tmrUnloader_Timer()
    tmrUnloader.Enabled = False
    Unload Me
End Sub

Public Sub WaitingOnTop(ByVal SetOnTop As Boolean)
    OnTop Me.hwnd, SetOnTop
End Sub

