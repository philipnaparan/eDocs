VERSION 5.00
Begin VB.Form frmReminders 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Reminders"
   ClientHeight    =   4530
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
   ScaleHeight     =   4530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Select which reminder you want to view: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   150
      TabIndex        =   2
      Top             =   750
      Width           =   4365
      Begin VB.OptionButton Option1 
         Caption         =   "Expired Alerts:"
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
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1890
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Alerts In Next 7 Days:"
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   600
         Width           =   2490
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Alerts In Next 15 Days:"
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Top             =   900
         Width           =   2490
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Alerts In Next 30 Days:"
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   1200
         Width           =   2490
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Alerts In Next 60 Days:"
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   1500
         Width           =   2490
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Alerts In Next 90 Days:"
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   1800
         Width           =   2490
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   375
         Width           =   1290
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   675
         Width           =   1290
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0"
         Top             =   975
         Width           =   1290
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1275
         Width           =   1290
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   1575
         Width           =   1290
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   1875
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   390
      Left            =   1875
      TabIndex        =   1
      Top             =   3900
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   390
      Left            =   3225
      TabIndex        =   0
      Top             =   3900
      Width           =   1230
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   -75
      TabIndex        =   16
      Top             =   600
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -1050
      TabIndex        =   17
      Top             =   3750
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   53
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   150
      Picture         =   "frmReminders.frx":0000
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reminders"
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
      TabIndex        =   15
      Top             =   150
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -150
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmReminders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    If Option1.Value = True Then LastViewedReminderOption = "exprd"
    If Option2.Value = True Then LastViewedReminderOption = "next7"
    If Option3.Value = True Then LastViewedReminderOption = "next15"
    If Option4.Value = True Then LastViewedReminderOption = "next30"
    If Option5.Value = True Then LastViewedReminderOption = "next60"
    If Option6.Value = True Then LastViewedReminderOption = "next90"
    
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LastViewedReminderOption = ""
    Dim fromDate As Date
    fromDate = DateAdd("d", 1, Date)
    
    Text1.Text = GetExpiredCount(CDate("1/1/1900"), Date)
    Text2.Text = GetExpiredCount(fromDate, DateAdd("d", 7, fromDate))
    Text3.Text = GetExpiredCount(fromDate, DateAdd("d", 15, fromDate))
    Text4.Text = GetExpiredCount(fromDate, DateAdd("d", 30, fromDate))
    Text5.Text = GetExpiredCount(fromDate, DateAdd("d", 60, fromDate))
    Text6.Text = GetExpiredCount(fromDate, DateAdd("d", 90, fromDate))
    
End Sub


Private Function GetExpiredCount(ByVal fromDate As Date, ByVal toDate As Date) As Long
    If AppDBType = adDBTypeSQLServer Then
        GetExpiredCount = GetRecordCount("SELECT [ID] FROM [vw_FileInfoOnly] WHERE [AlertDate] BETWEEN '" & Format$(fromDate, "yyyy/MM/dd") & "' AND '" & Format$(toDate, "yyyy/MM/dd") & "'")
    ElseIf AppDBType = adDBTypeMSAccess Then
        GetExpiredCount = GetRecordCount("SELECT [ID] FROM [vw_FileInfoOnly] WHERE [AlertDate] BETWEEN #" & Format$(fromDate, "yyyy/MM/dd") & "# AND #" & Format$(toDate, "yyyy/MM/dd") & "#")
    End If
End Function

