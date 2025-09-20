VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDUMMY 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin eDocs.ctrlWaiting ctrlWaiting3 
      Height          =   765
      Left            =   4500
      TabIndex        =   10
      Top             =   1725
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin eDocs.ctrlWaiting ctrlWaiting2 
      Height          =   765
      Left            =   1725
      TabIndex        =   9
      Top             =   1725
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin eDocs.ctrlWaiting ctrlWaiting1 
      Height          =   765
      Left            =   4500
      TabIndex        =   8
      Top             =   975
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.PictureBox pnlMetaData 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   1200
      ScaleHeight     =   4815
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   3225
      Width           =   4515
      Begin VB.CommandButton cmdDetailAdd 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   75
         TabIndex        =   6
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   5
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2325
         TabIndex        =   4
         Top             =   4350
         Width           =   1005
      End
      Begin VB.CommandButton cmdDetailRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   3
         Top             =   4350
         Width           =   1005
      End
      Begin MSComctlLib.ListView lstvRestrictions 
         Height          =   4140
         Left            =   -150
         TabIndex        =   7
         Top             =   600
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   7303
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group Name"
            Object.Width           =   72496
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   765
      Left            =   1275
      TabIndex        =   1
      Top             =   2625
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1349
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   1725
      TabIndex        =   0
      Top             =   975
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
End
Attribute VB_Name = "frmDUMMY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

