VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DD53C357-B171-4403-B656-34DFAB17A8B1}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User login"
   ClientHeight    =   2085
   ClientLeft      =   255
   ClientTop       =   2100
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckSavePass 
      Caption         =   "Save Password"
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   1170
      Width           =   1500
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   4875
      Top             =   1650
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   4
      Bmp:1           =   "frmLogin.frx":038A
      Key:1           =   "#mnuHAbout"
      Bmp:2           =   "frmLogin.frx":07B2
      Key:2           =   "#mnuHTT"
      Bmp:3           =   "frmLogin.frx":0BDA
      Key:3           =   "#mnuFCS"
      Bmp:4           =   "frmLogin.frx":1002
      Key:4           =   "#mnuFE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   1740
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
            Text            =   " Enter a valid username and password to login."
            TextSave        =   " Enter a valid username and password to login."
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   315
      Left            =   4170
      TabIndex        =   3
      Top             =   315
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4170
      TabIndex        =   4
      Top             =   765
      Width           =   1065
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   765
      Width           =   1965
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   2100
      MaxLength       =   20
      TabIndex        =   0
      Top             =   315
      Width           =   1965
   End
   Begin VB.PictureBox ctrlLiner2 
      BackColor       =   &H80000010&
      Height          =   10
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   6690
      TabIndex        =   8
      Top             =   0
      Width           =   6690
   End
   Begin eDocs.ctrlWaiting WaitingDisplay 
      Height          =   690
      Left            =   525
      TabIndex        =   9
      Top             =   3450
      Width           =   2640
      _ExtentX        =   4921
      _ExtentY        =   1349
   End
   Begin VB.Image imgUser 
      Height          =   960
      Left            =   150
      Picture         =   "frmLogin.frx":142A
      Top             =   150
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   765
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   315
      Width           =   990
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFCS 
         Caption         =   "&Connection Settings"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSqlSvrDB 
         Caption         =   "Use SQL Server Database"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFE 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHTT 
         Caption         =   "&Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuShowSplash 
         Caption         =   "&Show Splash Screen"
      End
      Begin VB.Menu mnuHSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About e-Docs..."
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    DoEvents
    
    If IsControlEmpty(txtUserName) Then Exit Sub
    If IsControlEmpty(txtPassword) Then Exit Sub
    
    If AppConnectionString = "" Then
        MsgBox "Connection is not yet configured!", vbCritical
        Exit Sub
    Else
        If IsDBValid(AppConnectionString) = 0 Then
            MsgBox "The database that you are connected is not valid for this application.", vbCritical
            mnuFCS_Click
            Exit Sub
        ElseIf IsDBValid(AppConnectionString) = -1 Then
            MsgBox "Unable to connect to '" & GetDataSource & "'.Please check the connection and try again!", vbCritical
            mnuFCS_Click
            Exit Sub
        End If
    End If
    
    Dim RsLogin As recordset
    
    Me.MousePointer = vbHourglass
    Set RsLogin = GetRecords("SELECT * FROM vw_Users " & _
                              "WHERE(" & _
                                    "(UserName='" & Encode(txtUserName.Text) & "') AND " & _
                                    "(UserPassword='" & Encode(txtPassword.Text) & "')" & _
                                    ")")
    
    If RsLogin Is Nothing Then
        MsgBox "Error retrieving user information." & vbCrLf & "Please check connection settings.", vbCritical, "Login Failed!"
    Else
        If RsLogin.RecordCount <> 0 Then
            With AppCurrentUser
                .UserName = txtUserName.Text
                .CompleteName = RsLogin.Fields("CompleteName")
                .UserGroupId = Val(RsLogin.Fields("UserGroupID"))
                .bCanViewConfidential = CBool(RsLogin.Fields("CanViewConfidential"))
                .bCanAdd = CBool(RsLogin.Fields("CanAdd"))
                .bCanEdit = CBool(RsLogin.Fields("CanEdit"))
                .bCanChkOut = CBool(RsLogin.Fields("CanCheckOut"))
                .bCanDelete = CBool(RsLogin.Fields("CanDelete"))
                .bCanImport = CBool(RsLogin.Fields("CanImport"))
                .bCanExport = CBool(RsLogin.Fields("CanExport"))
                .bCanAddFolder = CBool(RsLogin.Fields("CanAddFolder"))
                .bCanEditFolder = CBool(RsLogin.Fields("CanEditFolder"))
                .bCanDeleteFolder = CBool(RsLogin.Fields("CanDeleteFolder"))
                .bCanManageTemplates = CBool(RsLogin.Fields("CanManageTemplates"))
                .bIsSysAdmin = CBool(RsLogin.Fields("IsSystemAdministrator"))
                .UserId = RsLogin.Fields("ID")
            End With
            
            SaveLogInfo
            LoginOk = True
            
            
            Unload Me
        Else
            MsgBox "Invalid Username/Password.Please try again!", vbCritical, "Login Failed!"
            HighLightText txtUserName
            txtUserName.SetFocus
        End If
        
        If RsLogin.State = adStateOpen Then RsLogin.Close
        Set RsLogin = Nothing
    End If
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    GetLogInfo

    If mnuSqlSvrDB.Checked = True Then
        AppDBType = adDBTypeSQLServer
    Else
        AppDBType = adDBTypeMSAccess
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub mnuFCS_Click()
    On Error GoTo err
    
    Dim TestConn As ADODB.Connection
    Dim DBSource As MSDASC.DataLinks
    
    Set TestConn = New ADODB.Connection
    Set DBSource = New MSDASC.DataLinks
    
    DoEvents
       
    If AppConnectionString <> "" Then
        TestConn.ConnectionString = AppConnectionString
    End If
    
    DBSource.hwnd = Me.hwnd
    
    If DBSource.PromptEdit(TestConn) Then AppConnectionString = TestConn.ConnectionString

    Set TestConn = Nothing
    
    If IsDBValid(AppConnectionString) = 0 Then
            If MsgBox("The database that you are connected is not valid for this application." & vbCrLf & vbCrLf & _
                   "Would you like to setup the connection again?", vbCritical + vbYesNo) = vbYes Then
                
                mnuFCS_Click
                
            End If
        ElseIf IsDBValid(AppConnectionString) = -1 Then
            If MsgBox("Unable to connect to '" & GetDataSource & "'." & vbCrLf & vbCrLf & _
                   "Would you like to setup the connection again?", vbCritical + vbYesNo) = vbYes Then
                
                mnuFCS_Click
                
            End If
        End If
        
    Exit Sub
err:
    Select Case err.Number
        Case -2147217805
            TestConn.ConnectionString = ""
            Resume
    End Select
    
End Sub

Private Sub mnuFE_Click()
    Unload Me
End Sub

Private Sub SaveLogInfo()
    UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""dbconnection""]", "value", Encode(AppConnectionString), App.Path & "\settings\AppSettings.xml"
    UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""lastlogin""]", "value", Encode(txtUserName.Text), App.Path & "\settings\AppSettings.xml"
    UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""usesqlserver""]", "value", mnuSqlSvrDB.Checked, App.Path & "\settings\AppSettings.xml"
    
    If ckSavePass.Value = vbChecked Then
        UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""lastloginpwd""]", "value", Encode(txtPassword.Text), App.Path & "\settings\AppSettings.xml"
    Else
        UpdateXMLAttributeValue "/appsettings/section[@name=""system""]/item[@name=""lastloginpwd""]", "value", "", App.Path & "\settings\AppSettings.xml"
    End If
    
    If mnuSqlSvrDB.Checked = True Then
        AppDBType = adDBTypeSQLServer
    Else
        AppDBType = adDBTypeMSAccess
    End If
    
End Sub

Private Sub GetLogInfo()
    AppConnectionString = DeCode(GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""dbconnection""]", "value", App.Path & "\settings\AppSettings.xml"))
    txtUserName.Text = DeCode(GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""lastlogin""]", "value", App.Path & "\settings\AppSettings.xml"))
    txtPassword.Text = DeCode(GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""lastloginpwd""]", "value", App.Path & "\settings\AppSettings.xml"))
    mnuSqlSvrDB.Checked = CBool(GetXMLAttributeValue("/appsettings/section[@name=""system""]/item[@name=""usesqlserver""]", "value", App.Path & "\settings\AppSettings.xml"))
    
    If txtPassword.Text <> "" Then
        ckSavePass.Value = vbChecked
    Else
        ckSavePass.Value = vbUnchecked
    End If
    
End Sub

Private Sub mnuH_Click()
   ' Dim posRect As RECT
    
   ' GetClientRect Me.hwnd, posRect
    
    'PopupMenu MainForm.mnuH, , posRect.Left + 420, posRect.Top
End Sub

Private Sub mnuHAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHTT_Click()
    On Error Resume Next
    OpenURL App.Path & "\Help.doc", Me.hwnd
End Sub

Private Sub mnuShowSplash_Click()
    frmSplash.bForDisplay = True
    frmSplash.Show vbModal
End Sub

Private Sub mnuSqlSvrDB_Click()
    mnuSqlSvrDB.Checked = Not mnuSqlSvrDB.Checked
End Sub
