VERSION 5.00
Begin VB.Form frmTrialInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "e-Docs Trial"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrialInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReg 
      Caption         =   "&Register"
      Height          =   390
      Left            =   3000
      TabIndex        =   0
      Top             =   2325
      Width           =   1290
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   150
      TabIndex        =   5
      Top             =   2175
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4425
      TabIndex        =   1
      Top             =   2325
      Width           =   1215
   End
   Begin VB.Timer tmrAlert 
      Interval        =   500
      Left            =   4650
      Top             =   375
   End
   Begin VB.Label lblRemain 
      Caption         =   "DAYS REMAINING: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   525
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   $"frmTrialInfo.frx":038A
      Height          =   765
      Left            =   150
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Note:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   975
      Width           =   1365
   End
   Begin VB.Label lblAlert 
      Caption         =   "TRIAL WILL EXPIRE ON "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   5415
   End
End
Attribute VB_Name = "frmTrialInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim expiryDate As Date
Dim lastUseDate As Date

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    LastGenericText = "reg"
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    LastGenericText = ""
    
    Dim objExpiry As CExpiry
    Set objExpiry = New CExpiry
    
    lastUseDate = CDate(objExpiry.GetLastUse)
    expiryDate = CDate(objExpiry.GetExpiry)
    
    Set objExpiry = Nothing
    
    lblAlert.Caption = "TRIAL WILL EXPIRE ON  " & Format$(expiryDate, "MMM. dd, yyyy")
    lblAlert.Caption = UCase(lblAlert.Caption)
    
    
    lblRemain.Caption = "DAYS REMAINING:  " & DateDiff("d", lastUseDate, expiryDate)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTrialInfo = Nothing
End Sub

Private Sub tmrAlert_Timer()
    lblAlert.Visible = Not lblAlert.Visible
End Sub
