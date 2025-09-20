VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFolderRestrictionsAddEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolderRestrictionsAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Deny Open File"
      DataField       =   "DenyOpenFile"
      Height          =   240
      Left            =   1875
      TabIndex        =   1
      Top             =   1275
      Width           =   2190
   End
   Begin VB.CheckBox Check2 
      Caption         =   "DenyCreateFile"
      DataField       =   "DenyCreateFile"
      Height          =   240
      Left            =   1875
      TabIndex        =   2
      Top             =   1575
      Width           =   2190
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Deny Edit File"
      DataField       =   "DenyEditFile"
      Height          =   240
      Left            =   1875
      TabIndex        =   3
      Top             =   1875
      Width           =   2190
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Deny Delete File"
      CausesValidation=   0   'False
      DataField       =   "DenyDeleteFile"
      Height          =   240
      Left            =   1875
      TabIndex        =   4
      Top             =   2175
      Width           =   2190
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Deny Check Out"
      DataField       =   "DenyCheckOut"
      Height          =   240
      Left            =   1875
      TabIndex        =   5
      Top             =   2475
      Width           =   2190
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Deny Folder Delete"
      DataField       =   "DenyFolderDelete"
      Height          =   240
      Left            =   1875
      TabIndex        =   10
      Top             =   4050
      Width           =   2190
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Deny Folder Edit"
      DataField       =   "DenyFolderEdit"
      Height          =   240
      Left            =   1875
      TabIndex        =   9
      Top             =   3750
      Width           =   2190
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Deny Folder Access"
      DataField       =   "DenyFolderAccess"
      Height          =   240
      Left            =   1875
      TabIndex        =   8
      Top             =   3450
      Width           =   2190
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Deny File Import"
      DataField       =   "DenyFileImport"
      Height          =   240
      Left            =   1875
      TabIndex        =   6
      Top             =   2775
      Width           =   2190
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Deny File Export"
      DataField       =   "DenyFileExport"
      Height          =   240
      Left            =   1875
      TabIndex        =   7
      Top             =   3075
      Width           =   2190
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   3675
      TabIndex        =   12
      Top             =   4650
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   390
      Left            =   2325
      TabIndex        =   11
      Top             =   4650
      Width           =   1230
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   12915
      _extentx        =   22781
      _extenty        =   53
   End
   Begin eDocs.ctrlLiner ctrlLiner1 
      Height          =   30
      Index           =   0
      Left            =   -6900
      TabIndex        =   15
      Top             =   4500
      Width           =   13890
      _extentx        =   24500
      _extenty        =   53
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   750
      TabIndex        =   17
      Top             =   6000
      Width           =   2640
      _extentx        =   4921
      _extenty        =   1349
   End
   Begin MSDataListLib.DataCombo dcGroup 
      DataField       =   "UserGroupID"
      Height          =   315
      Left            =   1875
      TabIndex        =   0
      Top             =   825
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Files:"
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
      TabIndex        =   19
      Top             =   1275
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Folders:"
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
      TabIndex        =   18
      Top             =   3450
      Width           =   1590
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Group Name:"
      Height          =   315
      Left            =   225
      TabIndex        =   16
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
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
      Left            =   675
      TabIndex        =   14
      Top             =   150
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   225
      Picture         =   "frmFolderRestrictionsAddEdit.frx":038A
      Top             =   150
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -75
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmFolderRestrictionsAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lRecordPK As Long
Public lFolderId As Long

Dim bAddState As Boolean
Dim rsRecord As recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If IsControlEmpty(dcGroup) Then Exit Sub
    
    ShowWaiting Me, WaitingDisplay
    
    rsRecord.Fields("FolderId") = lFolderId
    rsRecord.Fields("UserGroupID") = dcGroup.BoundText
    
    SyncronizeRecordsetBinding rsRecord, Me
    SaveRecord "SELECT [FolderID],[UserGroupID],[DenyFolderAccess],[DenyFolderEdit],[DenyFolderDelete],[DenyOpenFile],[DenyCreateFile],[DenyEditFile],[DenyDeleteFile],[DenyCheckOut],[DenyFileImport],[DenyFileExport] FROM tbl_FolderRestrictions WHERE [ID]=" & lRecordPK, rsRecord, , Not bAddState
    
    HideWaiting Me, WaitingDisplay
    
    Unload Me
End Sub

Private Sub Form_Load()
    If lRecordPK = 0 Then bAddState = True

    ShowWaiting Me, WaitingDisplay
    
    Set rsRecord = GetRecords("SELECT * FROM tbl_FolderRestrictions WHERE [ID]=" & lRecordPK)
    If rsRecord Is Nothing Then
        MsgBox "Unable to read the data.", vbCritical, "Unexpected Error"
        Unload Me
        Exit Sub
    Else
        dcGroup.BoundColumn = "ID"
        dcGroup.ListField = "GroupName"
        Set dcGroup.RowSource = GetRecords("SELECT * FROM tbl_UserGroup WHERE [IsSystemDefined] IS NULL ORDER BY GroupName ASC")
        
        BindControlToRS Me, rsRecord
    End If
    
    If bAddState = True Then
        Me.Caption = "Add New Record"
        rsRecord.AddNew
        
        rsRecord.Fields("DenyFolderAccess") = 0
        rsRecord.Fields("DenyFolderEdit") = 0
        rsRecord.Fields("DenyFolderDelete") = 0
        rsRecord.Fields("DenyOpenFile") = 0
        rsRecord.Fields("DenyCreateFile") = 0
        rsRecord.Fields("DenyEditFile") = 0
        rsRecord.Fields("DenyDeleteFile") = 0
        rsRecord.Fields("DenyCheckOut") = 0
        rsRecord.Fields("DenyFileImport") = 0
        rsRecord.Fields("DenyFileExport") = 0
        
    Else
        Me.Caption = "Edit Record"
    End If
    
    
    
    lblTitle.Caption = Me.Caption
    
    HideWaiting Me, WaitingDisplay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFolderRestrictionsAddEdit = Nothing
End Sub

