VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   299
   ScaleMode       =   0  'User
   ScaleWidth      =   498.998
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5565
      Top             =   4095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public bForDisplay As Boolean

Private Sub Form_Click()
    If bForDisplay = True Then Unload Me
End Sub

Private Sub Form_Load()
    If bForDisplay = False Then trmUnload.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    bForDisplay = False
End Sub

Private Sub trmUnload_Timer()
    Unload Me
End Sub
