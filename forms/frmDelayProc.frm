VERSION 5.00
Begin VB.Form frmDelayProc 
   BorderStyle     =   0  'None
   ClientHeight    =   0
   ClientLeft      =   -4995
   ClientTop       =   -4995
   ClientWidth     =   0
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUnloader 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -450
      Top             =   -75
   End
End
Attribute VB_Name = "frmDelayProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    Me.Height = 0
    Me.Width = 0
    
    Me.Top = -4995
    Me.Left = -4995
End Sub

Private Sub Form_Load()
    tmrUnloader.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDelayProc = Nothing
End Sub

Private Sub tmrUnloader_Timer()
    tmrUnloader.Enabled = False
    Unload Me
End Sub
