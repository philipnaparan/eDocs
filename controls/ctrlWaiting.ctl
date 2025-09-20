VERSION 5.00
Begin VB.UserControl ctrlWaiting 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ScaleHeight     =   4635
   ScaleWidth      =   6180
   Begin VB.Timer tmrAnimateImage 
      Interval        =   100
      Left            =   2475
      Top             =   2100
   End
   Begin VB.Timer tmrAnim 
      Interval        =   500
      Left            =   1500
      Top             =   75
   End
   Begin VB.Timer tmrUnloader 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1725
      Top             =   825
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   11
      Left            =   2325
      Picture         =   "ctrlWaiting.ctx":0000
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   10
      Left            =   1950
      Picture         =   "ctrlWaiting.ctx":05C5
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   9
      Left            =   1575
      Picture         =   "ctrlWaiting.ctx":0B89
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   8
      Left            =   1200
      Picture         =   "ctrlWaiting.ctx":1140
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   7
      Left            =   825
      Picture         =   "ctrlWaiting.ctx":1704
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   6
      Left            =   450
      Picture         =   "ctrlWaiting.ctx":1CCC
      Top             =   3225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   5
      Left            =   2325
      Picture         =   "ctrlWaiting.ctx":2282
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   4
      Left            =   1950
      Picture         =   "ctrlWaiting.ctx":284A
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   3
      Left            =   1575
      Picture         =   "ctrlWaiting.ctx":2E0E
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   2
      Left            =   1200
      Picture         =   "ctrlWaiting.ctx":33CB
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   1
      Left            =   825
      Picture         =   "ctrlWaiting.ctx":3990
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAnimate 
      Height          =   480
      Left            =   75
      Top             =   150
      Width           =   480
   End
   Begin VB.Image imgAnimPic 
      Height          =   480
      Index           =   0
      Left            =   450
      Picture         =   "ctrlWaiting.ctx":3F50
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   1890
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
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   825
      TabIndex        =   0
      Top             =   150
      Width           =   1740
   End
End
Attribute VB_Name = "ctrlWaiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim I As Long
Dim lCurrentAnimImg As Long


Private Sub tmrAnim_Timer()
    On Error Resume Next
    I = I + 1
    If I > 3 Then I = 0
    
    If I = 0 Then
        lblWaitingMsg.Caption = "Processing"
    Else
        lblWaitingMsg.Caption = "Processing" & FillStr(".", I)
    End If
End Sub

Public Sub StartAnim()
    'tmrAnim.Enabled = True
End Sub

Private Sub tmrAnimateImage_Timer()
    On Error Resume Next
    lCurrentAnimImg = lCurrentAnimImg + 1
    If lCurrentAnimImg > 12 Then lCurrentAnimImg = 1
    
    imgAnimate.Picture = imgAnimPic(lCurrentAnimImg - 1).Picture
    
End Sub

Private Sub tmrUnloader_Timer()
    On Error Resume Next
    UserControl.Parent.WaitingDisplay.Visible = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Width = 2790
    UserControl.Height = 765
    
    shpBorder.Top = 0
    shpBorder.Left = 0
    
    shpBorder.Width = UserControl.Width
    shpBorder.Height = UserControl.Height
End Sub

Public Sub Terminate()
    On Error Resume Next
    'tmrAnim.Enabled = False
    tmrUnloader.Enabled = True
End Sub

