VERSION 5.00
Begin VB.Form frmProdActivation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Product Activation"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLic 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   75
      TabIndex        =   0
      Top             =   975
      Width           =   4515
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   75
      TabIndex        =   2
      Top             =   2475
      Width           =   4515
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1725
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3390
      TabIndex        =   4
      Top             =   3150
      Width           =   1140
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Default         =   -1  'True
      Height          =   390
      Left            =   2175
      TabIndex        =   3
      Top             =   3150
      Width           =   1140
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   6
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      Caption         =   "Licensed To:"
      Height          =   315
      Left            =   75
      TabIndex        =   10
      Top             =   750
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Activation Code:"
      Height          =   315
      Left            =   75
      TabIndex        =   9
      Top             =   2250
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Activation Key:"
      Height          =   315
      Left            =   75
      TabIndex        =   8
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Activation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   600
      TabIndex        =   7
      Top             =   150
      Width           =   3690
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmProdActivation.frx":0000
      Top             =   150
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmProdActivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim diskSerial As String

Private Sub cmdActivate_Click()
    If txtLic.Text = "" Then txtLic.SetFocus: Beep: Exit Sub
    If txtCode.Text = "" Then txtCode.SetFocus: Beep: Exit Sub
    
    LastGenericText = AESEncrypt(txtLic, EncPass, enum256Bit)
    ProductCode = txtCode.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LastGenericText = ""
    ProductCode = ""
    
    diskSerial = GetDriveSerial
    txtKey.Text = diskSerial
End Sub

Private Sub txtLic_LostFocus()
    txtLic.Text = UCase(txtLic.Text)
End Sub
