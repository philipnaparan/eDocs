VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Users (New)"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmUsers.frx":038A
      Left            =   1800
      List            =   "frmUsers.frx":0391
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   4185
      Width           =   3030
   End
   Begin VB.CommandButton cmdMail 
      Height          =   315
      Left            =   4875
      Picture         =   "frmUsers.frx":03A4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "E-mail"
      Top             =   3075
      Width           =   315
   End
   Begin VB.CheckBox ckAdmin 
      Caption         =   "Administrator"
      Height          =   315
      Left            =   4860
      TabIndex        =   9
      Top             =   3960
      Width           =   1290
   End
   Begin VB.TextBox txtLastUpdateBy 
      BackColor       =   &H00E6FFFF&
      Height          =   315
      Left            =   7575
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2025
      Width           =   2040
   End
   Begin VB.TextBox txtDateAdd 
      BackColor       =   &H00E6FFFF&
      Height          =   315
      Left            =   7575
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   900
      Width           =   2040
   End
   Begin VB.TextBox txtAddBy 
      BackColor       =   &H00E6FFFF&
      Height          =   315
      Left            =   7575
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2040
   End
   Begin VB.TextBox txtLastUpdate 
      BackColor       =   &H00E6FFFF&
      Height          =   315
      Left            =   7575
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1650
      Width           =   2040
   End
   Begin VB.CommandButton cmdFindCus 
      Caption         =   "&Find User"
      Height          =   315
      Left            =   75
      TabIndex        =   14
      Top             =   5025
      Width           =   1365
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   8475
      TabIndex        =   19
      Top             =   5025
      Width           =   1140
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   7275
      TabIndex        =   18
      Top             =   5025
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   4875
      TabIndex        =   16
      Top             =   5025
      Width           =   1140
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   315
      Left            =   6075
      TabIndex        =   17
      Top             =   5025
      Width           =   1140
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Create &New"
      Height          =   315
      Left            =   1500
      TabIndex        =   15
      Top             =   5025
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCPwd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   1800
      Width           =   1965
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   900
      Width           =   3015
   End
   Begin VB.TextBox txtNPwd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   1425
      Width           =   1965
   End
   Begin VB.TextBox txtFName 
      Height          =   315
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   3
      Top             =   2325
      Width           =   3015
   End
   Begin VB.TextBox txtLName 
      Height          =   315
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2700
      Width           =   3015
   End
   Begin VB.TextBox txtPhone 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   3450
      Width           =   1740
   End
   Begin VB.TextBox txtAPhone 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   3825
      Width           =   1740
   End
   Begin VB.TextBox txtEAdd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   3075
      Width           =   3015
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   9540
      TabIndex        =   21
      Top             =   4875
      Width           =   9540
   End
   Begin VB.Label Label7 
      Caption         =   "Rule"
      Height          =   315
      Left            =   150
      TabIndex        =   34
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmUsers.frx":092E
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label16 
      Caption         =   "Last Updated By"
      Height          =   315
      Left            =   6225
      TabIndex        =   33
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Label15 
      Caption         =   "Date Added"
      Height          =   315
      Left            =   6225
      TabIndex        =   32
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Label14 
      Caption         =   "Added By"
      Height          =   315
      Left            =   6225
      TabIndex        =   31
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Label13 
      Caption         =   "Last Updated"
      Height          =   315
      Left            =   6225
      TabIndex        =   30
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Password"
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
      TabIndex        =   29
      Top             =   1875
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Users"
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
      Height          =   390
      Left            =   600
      TabIndex        =   20
      Top             =   75
      Width           =   2640
   End
   Begin VB.Label Label3 
      Caption         =   "Username"
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
      TabIndex        =   28
      Top             =   900
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "New Password"
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
      TabIndex        =   27
      Top             =   1500
      Width           =   1365
   End
   Begin VB.Label Label5 
      Caption         =   "First Name"
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
      TabIndex        =   26
      Top             =   2325
      Width           =   1140
   End
   Begin VB.Label Label6 
      Caption         =   "Last Name"
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
      TabIndex        =   25
      Top             =   2700
      Width           =   1140
   End
   Begin VB.Label Label9 
      Caption         =   "Phone No."
      Height          =   315
      Left            =   150
      TabIndex        =   24
      Top             =   3450
      Width           =   1140
   End
   Begin VB.Label Label10 
      Caption         =   "Alt. Phone No."
      Height          =   315
      Left            =   150
      TabIndex        =   23
      Top             =   3825
      Width           =   1140
   End
   Begin VB.Label Label11 
      Caption         =   "E-mail Address"
      Height          =   315
      Left            =   150
      TabIndex        =   22
      Top             =   3075
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E1B391&
      BorderWidth     =   4
      X1              =   -600
      X2              =   10200
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   -225
      Picture         =   "frmUsers.frx":15F8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10380
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim conn                    As Connection
Dim rsUsers                 As Recordset

Public lPK                  As Long
Public bAddState            As Boolean

Private Sub InitRecordset()
    Set conn = New Connection
    Set rsUsers = New Recordset
    
    conn.ConnectionString = AppConnectionString
    conn.Open
    
    With rsUsers
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        Set .ActiveConnection = conn
    End With
    If bAddState = True Then
        rsUsers.Open "SELECT * FROM tbl_Users WHERE UserPK=" & lPK
    Else
        rsUsers.Open "SELECT * FROM qry_Users WHERE UserPK=" & lPK
    End If
    

    Set rsUsers.ActiveConnection = Nothing
    conn.Close
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If lPK = AppCurrentUser.UserPK Then
        MsgBox "You cannot delete your own account.", vbCritical
        Exit Sub
    End If

    If MsgBox("Delete current record?", vbCritical + vbYesNo) = vbNo Then Exit Sub

    Dim iDelResult As Long
    
    iDelResult = DeleteRecord("DELETE * FROM tbl_Users WHERE UserPK=" & lPK)

    If iDelResult = 0 Then
        MsgBox "Record has been deleted.", vbInformation, "Confirmed!"
        cmdNew_Click
    Else
        If iDelResult = -2147467259 Then
            MsgBox "Unable to delete because another record includes related records.", vbExclamation
        End If
    End If
End Sub

Private Sub cmdFindCus_Click()
    Dim ObjPK As New clsCustomString
    
    Set frmUsersFind.refPK = ObjPK
    frmUsersFind.Show vbModal
    
    If ObjPK.Text <> "" Then
        bAddState = False
        cmdDelete.Enabled = True
        cmdNew.Visible = True
        
        lPK = Val(ObjPK.Text)
        
        InitRecordset
        DisplayForEdit
        
        Label2.FontBold = False
        Label4.FontBold = False
        
        Me.Caption = "Manage Users (Alteration)"
    End If

    txtUserName.SetFocus
    Set ObjPK = Nothing
End Sub

Private Sub cmdMail_Click()
    If IsControlEmpty(txtEAdd) Then Exit Sub
    
    OpenURL "mailto:" & txtEAdd.Text, Me.hwnd

End Sub

Private Sub cmdNew_Click()
    bAddState = True
       
    CreateNewPK
    InitRecordset
    
    Label2.FontBold = True
    Label4.FontBold = True
    
    txtDateAdd.Text = ""
    txtAddBy.Text = ""
    txtLastUpdate.Text = ""
    txtLastUpdateBy.Text = ""
    
    ResetFields
    
    cmdNew.Visible = False
    cmdDelete.Enabled = False
    
    Me.Caption = "Manage Users (New)"
    
End Sub

Private Sub cmdReset_Click()
    If MsgBox("This will clear the input fields.Do you want to proceed?", vbExclamation + vbYesNo) = vbYes Then ResetFields
End Sub

Private Sub ResetFields()
    ClearTextBox Me
    ckAdmin.Value = 0
    
    txtUserName.SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Err1
    
    DoEvents
    
    
    
    If bAddState = True Then
        If IsControlEmpty(txtNPwd) Then Exit Sub
        If IsControlEmpty(txtCPwd) Then Exit Sub
        
        'Confirm password
        If Encode(txtNPwd.Text) <> Encode(txtCPwd.Text) Then
            MsgBox "Both password must be equal.Please confirm it and try again!", vbCritical
            txtNPwd.SetFocus
            Exit Sub
        End If
    Else
        
        'Check record concurrency for edit
        Dim iConcurrencyInfo As Integer
        
        iConcurrencyInfo = ConcurrencyInfo("SELECT * FROM tbl_Users WHERE UserPK=" & lPK, _
                                           Val(rsUsers.Fields("ConcurrencyId")))
    
        If iConcurrencyInfo = -1 Then
            MsgBox "Record is no longer exist because other user delete this record.", vbCritical
            
            cmdNew_Click
            
            Exit Sub
        Else
            If iConcurrencyInfo = 0 Then
                MsgBox "Unable to save the current record because another user commit some changes on it." & vbCrLf & _
                        "Click 'OK' to reload the record.", vbExclamation
                        
                InitRecordset
                DisplayForEdit
                
                Exit Sub
            End If
        End If
        
        'Confirm password
        If txtNPwd.Text <> "" Or txtCPwd.Text <> "" Then
            If Encode(txtNPwd.Text) <> Encode(txtCPwd.Text) Then
                MsgBox "Both password must be equal.Please confirm it and try again!", vbCritical
                txtNPwd.SetFocus
                Exit Sub
            End If
        End If
        
    End If
           
    conn.Open
    
    Set rsUsers.ActiveConnection = conn
    
    On Error GoTo Err2
    
    With rsUsers
        If bAddState = True Then
            .AddNew
            
            .Fields("UserPK") = lPK
            .Fields("DateAdded") = Now
            .Fields("AddedByIdFK") = AppCurrentUser.UserPK
            
            .Fields("UserPassword") = Encode(txtNPwd.Text)
        Else
            .Fields("DateModified") = Now
            .Fields("LastUpdatedByFK") = AppCurrentUser.UserPK
            
            txtLastUpdate = Format$(.Fields("DateModified"), "MMM-dd-yyyy hh:mm:ss AM/PM")
            txtLastUpdateBy = AppCurrentUser.UserName
            
            If txtNPwd.Text <> "" Then .Fields("UserPassword") = Encode(txtNPwd.Text)
        End If
    
        .Fields("UserName") = Encode(txtUserName.Text)
        .Fields("UserNameD") = txtUserName.Text
        
        .Fields("FirstName") = txtFName.Text
        .Fields("LastName") = txtLName.Text
        .Fields("Email") = txtEAdd.Text
        .Fields("PhoneNo") = txtPhone.Text
        .Fields("AltPhoneNo") = txtAPhone.Text
        .Fields("IsAdmin") = ckAdmin.Value
        .Fields("ConcurrencyId") = GetConcurrencyId("tbl_Users")
        
        .Update
    End With
    
    
    Set rsUsers.ActiveConnection = Nothing
    conn.Close
       
    If bAddState = True Then
        CreateNewPK
        InitRecordset
        
        txtDateAdd.Text = ""
        txtAddBy.Text = ""
        txtLastUpdate.Text = ""
        txtLastUpdateBy.Text = ""
        
        ResetFields
        
        MsgBox "New record has been added.", vbInformation, "Confirmed!"
    Else
        MsgBox "Changes has been saved.", vbInformation, "Confirmed!"
        MainForm.CheckUserAccess
        MainForm.lblUserInfo.Caption = "Welcome, " & AppCurrentUser.UserName
    End If
    
    
    Exit Sub
Err1:
    Resume Next
Err2:
    MsgBox err.Description, vbCritical, err.Number
End Sub


Private Sub Form_Load()
    If bAddState = True Then
        CreateNewPK
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    
    InitRecordset
End Sub

Private Sub CreateNewPK()
    lPK = GetRowPK("tbl_Users")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set conn = Nothing
    Set rsUsers = Nothing
    
    Set frmUsers = Nothing
End Sub

Private Sub DisplayForEdit()
    On Error GoTo err
    
    With rsUsers
    
         txtDateAdd = Format$(.Fields("DateAdded"), "MMM-dd-yyyy hh:mm:ss AM/PM")
         txtAddBy = .Fields("AddedBy")
         txtLastUpdate = Format$(.Fields("DateModified"), "MMM-dd-yyyy hh:mm:ss AM/PM")
         txtLastUpdateBy = .Fields("LastUpdatedBy")
    
         txtUserName.Text = .Fields("UserNameD")
         txtFName.Text = .Fields("FirstName")
         txtLName.Text = .Fields("LastName")
         txtEAdd.Text = .Fields("Email")
         txtPhone.Text = .Fields("PhoneNo")
         txtAPhone.Text = .Fields("AltPhoneNo")
         If .Fields("IsAdmin") Then
            ckAdmin.Value = 1
         Else
            ckAdmin.Value = 0
        End If
        
    End With
    
    Exit Sub
err:
    If err.Number = 94 Then Resume Next
End Sub
