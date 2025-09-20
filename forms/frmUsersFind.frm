VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUsersFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find User"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsersFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOperation 
      Height          =   315
      ItemData        =   "frmUsersFind.frx":038A
      Left            =   1800
      List            =   "frmUsersFind.frx":03A3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "LIKE"
      Top             =   1350
      Width           =   2250
   End
   Begin VB.ComboBox cmbFields 
      Height          =   315
      ItemData        =   "frmUsersFind.frx":0422
      Left            =   1800
      List            =   "frmUsersFind.frx":0432
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   975
      Width           =   3420
   End
   Begin VB.TextBox txtLookFor 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1725
      Width           =   3390
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   975
      Width           =   1140
   End
   Begin VB.CommandButton cmdDisplayAll 
      Caption         =   "&Display All"
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      Top             =   1350
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cboSortBy 
      Height          =   315
      ItemData        =   "frmUsersFind.frx":045F
      Left            =   1800
      List            =   "frmUsersFind.frx":046F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2100
      Width           =   2265
   End
   Begin VB.ComboBox cboSortType 
      Height          =   315
      ItemData        =   "frmUsersFind.frx":049C
      Left            =   4125
      List            =   "frmUsersFind.frx":04A6
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2100
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar prog 
      Height          =   315
      Left            =   975
      TabIndex        =   18
      Top             =   3900
      Visible         =   0   'False
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5400
      TabIndex        =   11
      Top             =   5475
      Width           =   1140
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   5475
      Width           =   1140
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   1425
      ScaleHeight     =   30
      ScaleWidth      =   5115
      TabIndex        =   13
      Top             =   2625
      Width           =   5115
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2565
      Left            =   150
      TabIndex        =   9
      Top             =   2775
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   4524
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   7323903
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   1575
      ScaleHeight     =   30
      ScaleWidth      =   4965
      TabIndex        =   15
      Top             =   825
      Width           =   4965
   End
   Begin MSComCtl2.DTPicker dtpDate1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1725
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20709379
      CurrentDate     =   38207
   End
   Begin MSComCtl2.DTPicker dtpDate2 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   1725
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20709379
      CurrentDate     =   38207
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "And"
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Keywords:"
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
      Left            =   150
      TabIndex        =   22
      Top             =   1725
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Records Where?"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   21
      Top             =   975
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Condition:"
      Height          =   315
      Left            =   150
      TabIndex        =   20
      Top             =   1350
      Width           =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "Sort By:"
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   2100
      Width           =   1440
   End
   Begin VB.Label lblLoadingInfo 
      Height          =   315
      Left            =   150
      TabIndex        =   17
      Top             =   5475
      Width           =   3990
   End
   Begin VB.Label Label4 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   750
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmUsersFind.frx":04B5
      Top             =   150
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E1B391&
      BorderWidth     =   4
      X1              =   -600
      X2              =   10200
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCF3ED&
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   14
      Top             =   75
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "Search Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   2550
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   -225
      Picture         =   "frmUsersFind.frx":117F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10380
   End
End
Attribute VB_Name = "frmUsersFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub InitGrid()

    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 8
        .ColSel = 7
        'Initialize the column size
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 0
        
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "User Name"
        .TextMatrix(0, 2) = "First Name"
        .TextMatrix(0, 3) = "Last Name"
        .TextMatrix(0, 4) = "Email"
        .TextMatrix(0, 5) = "Phone No"
        .TextMatrix(0, 6) = "Alt. Phone No"
        .TextMatrix(0, 7) = "PK"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbLeftJustify
        .ColAlignment(7) = vbLeftJustify
    End With
    
End Sub

Public Sub DisplayResult()
    Dim rsResult As Recordset
    Dim strFilter As String
    
    Select Case cmbOperation.ListIndex
        Case 0: strFilter = " LIKE '%" & txtLookFor.Tag & "%' "
        Case 1: strFilter = " = '" & txtLookFor.Tag & "' "
        Case 2: strFilter = " <> '" & txtLookFor.Tag & "' "
        Case 3: strFilter = " > '" & txtLookFor.Tag & "' "
        Case 4: strFilter = " >= '" & txtLookFor.Tag & "' "
        Case 5: strFilter = " < '" & txtLookFor.Tag & "' "
        Case 6: strFilter = " <= '" & txtLookFor.Tag & "' "
        Case 7: strFilter = " BETWEEN #" & Format$(dtpDate1.Value, "yyyy/MM/dd") & "# AND #" & Format$(dtpDate2.Value, "yyyy/MM/dd") & "# "
    End Select
    
    If cmbFields.ListIndex = 0 Then
        strFilter = "(" & Replace(cmbFields.Text, " ", "") & "D " & strFilter & ")"
    Else
        strFilter = "(" & Replace(cmbFields.Text, " ", "") & " " & strFilter & ")"
    End If

    If txtLookFor.Text <> "" Then
        'Display only those records match for the search criteria
        Set rsResult = GetRecords("SELECT * FROM tbl_Users " & _
                                  "WHERE(" & strFilter & _
                                        ") " & _
                                   "ORDER BY " & Replace(cboSortBy.Text, " ", "") & " " & cboSortType.Text, True)
                                    
    Else
        'Display all records
        Set rsResult = GetRecords("SELECT * FROM tbl_Users ORDER BY " & Replace(cboSortBy.Text, " ", "") & " " & cboSortType.Text, True)
    End If
    
    'On Error GoTo err
    If rsResult Is Nothing Then
        MsgBox "Error retrieving records." & vbCrLf & "Please check connection settings.", vbCritical, "Login Failed!"
    Else
    
        If rsResult.RecordCount > 0 Then
            
            prog.Visible = True
            
            With rsResult
                If .RecordCount > 0 Then
                
                    prog.Min = 0
                    prog.Max = .RecordCount
                    
                    .MoveFirst
                    
                    Do
                        DoEvents
                        With Grid
                            If .Rows = 2 And .TextMatrix(1, 7) = "" Then
                                .TextMatrix(1, 1) = rsResult.Fields("UserNameD")
                                .TextMatrix(1, 2) = rsResult.Fields("FirstName")
                                .TextMatrix(1, 3) = rsResult.Fields("LastName")
                                .TextMatrix(1, 4) = rsResult.Fields("Email")
                                .TextMatrix(1, 5) = rsResult.Fields("PhoneNo")
                                .TextMatrix(1, 6) = rsResult.Fields("AltPhoneNo")
                                .TextMatrix(1, 7) = rsResult.Fields("UserPK")
                            Else
                                .Rows = .Rows + 1
                                .TextMatrix(.Rows - 1, 1) = rsResult.Fields("UserNameD")
                                .TextMatrix(.Rows - 1, 2) = rsResult.Fields("FirstName")
                                .TextMatrix(.Rows - 1, 3) = rsResult.Fields("LastName")
                                .TextMatrix(.Rows - 1, 4) = rsResult.Fields("Email")
                                .TextMatrix(.Rows - 1, 5) = rsResult.Fields("PhoneNo")
                                .TextMatrix(.Rows - 1, 6) = rsResult.Fields("AltPhoneNo")
                                .TextMatrix(.Rows - 1, 7) = rsResult.Fields("UserPK")
                            End If
                        End With
                        
                        lblLoadingInfo.Caption = "Loading " & .AbsolutePosition & " of " & .RecordCount
                        prog.Value = .AbsolutePosition
                        
                        .MoveNext
                    Loop While Not .EOF
                    
                Grid.Row = 1
                Grid.ColSel = 7
                
                End If
            End With
            
            lblLoadingInfo.Caption = "Total Record: " & rsResult.RecordCount
            prog.Visible = False
            
            MsgBox "There are " & rsResult.RecordCount & " record(s) found.", vbInformation
            
            cmdSelect.Enabled = True
            cmdSelect.Default = True
            
            Grid.SetFocus
        Else
            MsgBox "No record found.", vbInformation
            cmdSelect.Enabled = False
        End If
        
        If rsResult.State = adStateOpen Then rsResult.Close
        Set rsResult = Nothing
        
    End If
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmbOperation_Click()
    If cmbOperation.ListIndex = 7 Then
            dtpDate1.Visible = True
            dtpDate2.Visible = True
            txtLookFor.Visible = False
            
            dtpDate1.Value = Date
            dtpDate2.Value = Date
        Else
            txtLookFor.Visible = True
            dtpDate1.Visible = False
            dtpDate2.Visible = False
    End If
    
    Select Case cmbOperation.ListIndex
        Case 0: cmbOperation.Tag = "LIKE"
        Case 1: cmbOperation.Tag = "="
        Case 2: cmbOperation.Tag = "<>"
        Case 3: cmbOperation.Tag = ">"
        Case 4: cmbOperation.Tag = ">="
        Case 5: cmbOperation.Tag = "<"
        Case 6: cmbOperation.Tag = "<="
        Case 7: cmbOperation.Tag = "BETWEEN"
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDisplayAll_Click()
    DoEvents
       
    InitGrid
    DisplayResult
End Sub

Private Sub cmdSearch_Click()
    DoEvents
    
    If IsControlEmpty(txtLookFor) Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    InitGrid
    DisplayResult
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click()
    refPK.Text = Grid.TextMatrix(Grid.Row, 7)
    Unload Me
End Sub

Private Sub Form_Load()
    InitGrid
    cboSortBy.ListIndex = 0
    cboSortType.ListIndex = 0
    cmbFields.ListIndex = 0
    cmbOperation.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmUsersFind = Nothing
End Sub

Private Sub Grid_DblClick()
    If cmdSelect.Enabled = True Then cmdSelect_Click
End Sub

Private Sub txtLookFor_Change()
    txtLookFor.Tag = Replace(txtLookFor.Text, " ", "")
    txtLookFor.Tag = Replace(txtLookFor.Text, "/", "")
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Grid_DblClick
End Sub


