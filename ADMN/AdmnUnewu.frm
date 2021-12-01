VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form AdmnUnewu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add A New User"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optFrm 
      Caption         =   "From New"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add"
      Height          =   300
      Left            =   5280
      TabIndex        =   17
      ToolTipText     =   "Add This New User"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtNik 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "2"
      ToolTipText     =   "Required-Not Case Sensitive (20 max)"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtInt 
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Required-Not Case Sensitive (3 max)"
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox cmbGrp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Tag             =   "8"
      ToolTipText     =   "Required-Select User Class From List"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox optAct 
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   2880
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtNme 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Required-Not Case Sensitive (40 max)"
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtUid 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Required-Not Case Sensitive (30 max)"
      Top             =   600
      Width           =   2085
   End
   Begin VB.TextBox txtPas 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Case Sensitive Max (15) Char"
      Top             =   960
      Width           =   1600
   End
   Begin VB.TextBox txtVpw 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Verify Password"
      Top             =   1320
      Width           =   1600
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3480
      FormDesignWidth =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Areas Must Be Filled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   18
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Initials"
      Height          =   255
      Index           =   6
      Left            =   4080
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Class"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Active User?"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "AdmnUnewu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/9/05 Corrected opening and closing files
Dim bOnLoad As Byte
Dim iFreeFile1 As Integer
Dim iFreeFile2 As Integer

Dim bUserInstalled As Boolean

Dim sPassword As String
Dim sVerify As String
Dim sUserLcId As String
Dim sUserUcId As String
Dim sUserName As String
Dim sUserNick As String
Dim sUserInit As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub OpenDbfFiles()
   On Error Resume Next
   Close iFreeIdx
   Close iFreeDbf
   iFreeIdx = FreeFile
   Open sFilePath & "rstval.eid" For Random Shared As iFreeIdx Len = Len(SecPw)
   iFreeDbf = FreeFile
   Open sFilePath & "rstval.edd" For Random Shared As iFreeDbf Len = Len(Secure)
   
End Sub

Private Sub cmbGrp_Click()
   If cmbGrp.ListIndex = 0 Then
      lblClass = "Sets All User Permissions"
   Else
      lblClass = "Sets No User Permissions"
   End If
   
End Sub

Private Sub cmbGrp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   cmbGrp = Trim(cmbGrp)
   For iList = 0 To cmbGrp.ListCount - 1
      If cmbGrp = cmbGrp.List(iList) Then b = 1
   Next
   If b = 0 Then
      'Beep
      cmbGrp = cmbGrp.List(0)
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdNew_Click()
   Dim b As Byte
   sUserLcId = Trim(txtUid)
   sUserUcId = UCase$(sUserLcId)
   sUserName = Trim(txtNme)
   sUserNick = Trim(txtNik)
   sUserInit = Trim(txtInt)
   On Error Resume Next
   If sUserLcId = "" Or sUserName = "" Or sUserNick = "" Or sUserInit = "" Then
      MsgBox "One Or More Text Boxes Are Empty.", _
         vbInformation, Caption
      txtNme.SetFocus
      Exit Sub
   End If
   sPassword = Trim(txtPas)
   sVerify = Trim(txtVpw)
   If Len(sPassword) < 3 Then
      MsgBox "The Password Must Be At Least (3) Characters.", _
         vbInformation, Caption
      txtPas.SetFocus
      Exit Sub
   End If
   
   If sPassword <> sVerify Then
      MsgBox "The Password Must Be The Same As The Verification.", _
         vbInformation, Caption
      txtPas.SetFocus
      Exit Sub
   End If
   
   b = VerifyUserId()
   If b = 1 Then
      MsgBox "The User Id Is In Use (May Be A Reserved Id)" & vbCr _
         & "Please Select Another Unique User Id..", _
         vbInformation, Caption
   Else
      InstallUser
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      If optFrm.value = vbUnchecked Then cmbGrp.Enabled = False
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move MDISect.Left + 700, MDISect.Top + 1500
   FormatControls
   OpenDbfFiles
   cmbGrp.AddItem "Administrators"
   cmbGrp.AddItem "Users"
   cmbGrp = cmbGrp.List(0)
   lblClass = "Sets All User Permissions"
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   If optFrm.value = vbChecked And bUserInstalled Then
      'From Users
      AdmnUuser2.optFrm.value = vbChecked
   Else
      If bUserInstalled Then
         'First time
         iUserIdx = LOF(iFreeIdx) \ Len(SecPw)
         sSql = "UPDATE Preferences SET PreSec=1"
         clsADOCon.ExecuteSQL sSql
         bSecSet = 1
         SaveSetting "Esi2000", "System", "UserProfileRec", iUserIdx
         MsgBox "Congratulations. Advanced Security Is " & vbCr _
            & "Installed.  You May Now Install Other Users.", _
            vbInformation, Caption
      Else
         If optFrm.value = vbUnchecked Then
            MsgBox "Advanced Security Was Not Installed.", _
               vbExclamation, Caption
         End If
      End If
   End If
   If iUserIdx > 0 Then
      bSecSet = 1
      Get #iFreeIdx, iUserIdx, SecPw
      Get #iFreeDbf, iUserIdx, Secure
      SaveSetting "Esi2000", "System", "UserProfileRec", iUserIdx
   End If
   AdmnUuser2.FormatGrid
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set AdmnUnewu = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtPas = String(15, " ")
   txtVpw = String(15, " ")
   
End Sub


Private Sub txtInt_LostFocus()
   txtInt = CheckLen(txtInt, 3)
   
End Sub


Private Sub txtNik_LostFocus()
   txtNik = CheckLen(txtNik, 20)
   
End Sub


Private Sub txtNme_LostFocus()
   txtNme = CheckLen(txtNme, 40)
   txtNme = StrCase(txtNme)
   
End Sub


Private Sub txtPas_LostFocus()
   txtPas = CheckLen(txtPas, 15)
   If txtPas = "" Then txtPas = String(15, " ")
   
End Sub


Private Sub txtUid_LostFocus()
   txtUid = CheckLen(txtUid, 30)
   
End Sub


Private Sub txtVpw_LostFocus()
   txtVpw = CheckLen(txtVpw, 15)
   If txtVpw = "" Then txtVpw = String(15, " ")
   
End Sub



Private Function VerifyUserId() As Byte
   Dim iList As Integer
   Dim iOldrec As Integer
   Dim iLen As Integer
   iOldrec = iCurrentRec
   
   VerifyUserId = 0
   iLen = LOF(iFreeIdx) \ Len(SecPw)
   For iList = 1 To iLen
      Get #iFreeIdx, iList, SecPw
      If sUserUcId = Trim(SecPw.UserUcName) Then
         VerifyUserId = 1
         Exit For
      End If
   Next
   iCurrentRec = iOldrec
   
End Function

Private Sub InstallUser()
   Dim b As Byte
   Dim iList As Integer
   Dim iLen As Integer
   Dim iOldrec As Integer
   
   On Error Resume Next
   iOldrec = iCurrentRec
   If Left(cmbGrp, 1) = "A" Then b = 1 Else b = 0
   iList = (LOF(iFreeIdx) \ Len(SecPw)) + 1
   sPassword = ScramblePw(sPassword)
   SecPw.PassWord = sPassword
   SecPw.UserLcName = sUserLcId
   SecPw.UserUcName = sUserUcId
   SecPw.UserAdmn = b
   SecPw.UserRecord = iList
   
   Secure.UserName = sUserName
   Secure.UserInitials = sUserInit
   Secure.UserNickName = sUserNick
   
   Secure.UserActive = 1
   Secure.UserAddUser = 1
   If SecPw.UserAdmn = 1 Then
      Secure.UserLevel = 10
   Else
      Secure.UserLevel = 20
   End If
   Secure.UserNew = Format$(Now, "mm/dd/yy")
   InitializePermissions Secure, b
   '    Secure.UserAdmn = b
   '    Secure.UserAdmnG1 = b
   '    Secure.UserAdmnG1E = b
   '    Secure.UserAdmnG1V = b
   '    Secure.UserAdmnG1F = b
   '
   '    Secure.UserAdmnG2 = b
   '    Secure.UserAdmnG2E = b
   '    Secure.UserAdmnG2V = b
   '    Secure.UserAdmnG2F = b
   '
   '    Secure.UserAdmnG3 = b
   '    Secure.UserAdmnG3E = b
   '    Secure.UserAdmnG3V = b
   '    Secure.UserAdmnG3F = b
   '
   '    Secure.UserAdmnG4 = b
   '    Secure.UserAdmnG4E = b
   '    Secure.UserAdmnG4V = b
   '    Secure.UserAdmnG4F = b
   '
   '    Secure.UserAdmnG5 = b
   '    Secure.UserAdmnG5E = b
   '    Secure.UserAdmnG5V = b
   '    Secure.UserAdmnG5F = b
   '
   '    Secure.UserAdmnG6 = b
   '    Secure.UserAdmnG6E = b
   '    Secure.UserAdmnG6V = b
   '    Secure.UserAdmnG6F = b
   '
   '    Secure.UserAdmnG7 = b
   '    Secure.UserAdmnG7E = b
   '    Secure.UserAdmnG7V = b
   '    Secure.UserAdmnG7F = b
   '
   '    Secure.UserAdmnG8 = b
   '    Secure.UserAdmnG8E = b
   '    Secure.UserAdmnG8V = b
   '    Secure.UserAdmnG8F = b
   '
   '    Secure.UserEngr = b
   '    Secure.UserEngrG1 = b
   '    Secure.UserEngrG1E = b
   '    Secure.UserEngrG1V = b
   '    Secure.UserEngrG1F = b
   '
   '    Secure.UserEngrG2 = b
   '    Secure.UserEngrG2E = b
   '    Secure.UserEngrG2V = b
   '    Secure.UserEngrG2F = b
   '
   '    Secure.UserEngrG3 = b
   '    Secure.UserEngrG3E = b
   '    Secure.UserEngrG3V = b
   '    Secure.UserEngrG3F = b
   '
   '    Secure.UserEngrG4 = b
   '    Secure.UserEngrG4E = b
   '    Secure.UserEngrG4V = b
   '    Secure.UserEngrG4F = b
   '
   '    Secure.UserEngrG5 = b
   '    Secure.UserEngrG5E = b
   '    Secure.UserEngrG5V = b
   '    Secure.UserEngrG5F = b
   '
   '    Secure.UserEngrG6 = b
   '    Secure.UserEngrG6E = b
   '    Secure.UserEngrG6V = b
   '    Secure.UserEngrG6F = b
   '
   '    Secure.UserEngrG7 = b
   '    Secure.UserEngrG7E = b
   '    Secure.UserEngrG7V = b
   '    Secure.UserEngrG7F = b
   '
   '    Secure.UserEngrG8 = b
   '    Secure.UserEngrG8E = b
   '    Secure.UserEngrG8V = b
   '    Secure.UserEngrG8F = b
   '
   '    Secure.UserFina = b
   '    Secure.UserFinaG1 = b
   '    Secure.UserFinaG1E = b
   '    Secure.UserFinaG1V = b
   '    Secure.UserFinaG1F = b
   '
   '    Secure.UserFinaG2 = b
   '    Secure.UserFinaG2E = b
   '    Secure.UserFinaG2V = b
   '    Secure.UserFinaG2F = b
   '
   '    Secure.UserFinaG3 = b
   '    Secure.UserFinaG3E = b
   '    Secure.UserFinaG3V = b
   '    Secure.UserFinaG3F = b
   '
   '    Secure.UserFinaG4 = b
   '    Secure.UserFinaG4E = b
   '    Secure.UserFinaG4V = b
   '    Secure.UserFinaG4F = b
   '
   '    Secure.UserFinaG5 = b
   '    Secure.UserFinaG5E = b
   '    Secure.UserFinaG5V = b
   '    Secure.UserFinaG5F = b
   '
   '    Secure.UserFinaG6 = b
   '    Secure.UserFinaG6E = b
   '    Secure.UserFinaG6V = b
   '    Secure.UserFinaG6F = b
   '
   '    Secure.UserFinaG7 = b
   '    Secure.UserFinaG7E = b
   '    Secure.UserFinaG7V = b
   '    Secure.UserFinaG7F = b
   '
   '    Secure.UserFinaG8 = b
   '    Secure.UserFinaG8E = b
   '    Secure.UserFinaG8V = b
   '    Secure.UserFinaG8F = b
   '
   '    Secure.UserFinaG9 = b
   '    Secure.UserFinaG9E = b
   '    Secure.UserFinaG9V = b
   '    Secure.UserFinaG9F = b
   '
   '    Secure.UserFinaG10 = b
   '    Secure.UserFinaG10E = b
   '    Secure.UserFinaG10V = b
   '    Secure.UserFinaG10F = b
   '
   '
   '    Secure.UserInvc = b
   '    Secure.UserInvcG1 = b
   '    Secure.UserInvcG1E = b
   '    Secure.UserInvcG1V = b
   '    Secure.UserInvcG1F = b
   '
   '    Secure.UserInvcG2 = b
   '    Secure.UserInvcG2E = b
   '    Secure.UserInvcG2V = b
   '    Secure.UserInvcG2F = b
   '
   '    Secure.UserInvcG3 = b
   '    Secure.UserInvcG3E = b
   '    Secure.UserInvcG3V = b
   '    Secure.UserInvcG3F = b
   '
   '    Secure.UserInvcG4 = b
   '    Secure.UserInvcG4E = b
   '    Secure.UserInvcG4V = b
   '    Secure.UserInvcG4F = b
   '
   '    Secure.UserInvcG5 = b
   '    Secure.UserInvcG5E = b
   '    Secure.UserInvcG5V = b
   '    Secure.UserInvcG5F = b
   '
   '    Secure.UserInvcG6 = b
   '    Secure.UserInvcG6E = b
   '    Secure.UserInvcG6V = b
   '    Secure.UserInvcG6F = b
   '
   '    Secure.UserInvcG7 = b
   '    Secure.UserInvcG7E = b
   '    Secure.UserInvcG7V = b
   '    Secure.UserInvcG7F = b
   '
   '    Secure.UserInvcG8 = b
   '    Secure.UserInvcG8E = b
   '    Secure.UserInvcG8V = b
   '    Secure.UserInvcG8F = b
   '
   '    Secure.UserInvcG9 = b
   '    Secure.UserInvcG9E = b
   '    Secure.UserInvcG9V = b
   '    Secure.UserInvcG9F = b
   '
   '    Secure.UserInvcG10 = b
   '    Secure.UserInvcG10E = b
   '    Secure.UserInvcG10V = b
   '    Secure.UserInvcG10F = b
   '
   '
   '    Secure.UserProd = b
   '    Secure.UserProdG1 = b
   '    Secure.UserProdG1E = b
   '    Secure.UserProdG1V = b
   '    Secure.UserProdG1F = b
   '
   '    Secure.UserProdG2 = b
   '    Secure.UserProdG2E = b
   '    Secure.UserProdG2V = b
   '    Secure.UserProdG2F = b
   '
   '    Secure.UserProdG3 = b
   '    Secure.UserProdG3E = b
   '    Secure.UserProdG3V = b
   '    Secure.UserProdG3F = b
   '
   '    Secure.UserProdG4 = b
   '    Secure.UserProdG4E = b
   '    Secure.UserProdG4V = b
   '    Secure.UserProdG4F = b
   '
   '    Secure.UserProdG5 = b
   '    Secure.UserProdG5E = b
   '    Secure.UserProdG5V = b
   '    Secure.UserProdG5F = b
   '
   '    Secure.UserProdG6 = b
   '    Secure.UserProdG6E = b
   '    Secure.UserProdG6V = b
   '    Secure.UserProdG6F = b
   '
   '    Secure.UserProdG7 = b
   '    Secure.UserProdG7E = b
   '    Secure.UserProdG7V = b
   '    Secure.UserProdG7F = b
   '
   '    Secure.UserProdG8 = b
   '    Secure.UserProdG8E = b
   '    Secure.UserProdG8V = b
   '    Secure.UserProdG8F = b
   '
   '    Secure.UserProdG9 = b
   '    Secure.UserProdG9E = b
   '    Secure.UserProdG9V = b
   '    Secure.UserProdG9F = b
   '
   '    Secure.UserFinaG10 = b
   '    Secure.UserFinaG10E = b
   '    Secure.UserFinaG10V = b
   '    Secure.UserFinaG10F = b
   '
   '    Secure.UserQual = b
   '    Secure.UserQualG1 = b
   '    Secure.UserQualG1E = b
   '    Secure.UserQualG1V = b
   '    Secure.UserQualG1F = b
   '
   '    Secure.UserQualG2 = b
   '    Secure.UserQualG2E = b
   '    Secure.UserQualG2V = b
   '    Secure.UserQualG2F = b
   '
   '    Secure.UserQualG3 = b
   '    Secure.UserQualG3E = b
   '    Secure.UserQualG3V = b
   '    Secure.UserQualG3F = b
   '
   '    Secure.UserQualG4 = b
   '    Secure.UserQualG4E = b
   '    Secure.UserQualG4V = b
   '    Secure.UserQualG4F = b
   '
   '    Secure.UserQualG5 = b
   '    Secure.UserQualG5E = b
   '    Secure.UserQualG5V = b
   '    Secure.UserQualG5F = b
   '
   '    Secure.UserQualG6 = b
   '    Secure.UserQualG6E = b
   '    Secure.UserQualG6V = b
   '    Secure.UserQualG6F = b
   '
   '    Secure.UserQualG7 = b
   '    Secure.UserQualG7E = b
   '    Secure.UserQualG7V = b
   '    Secure.UserQualG7F = b
   '
   '    Secure.UserQualG8 = b
   '    Secure.UserQualG8E = b
   '    Secure.UserQualG8V = b
   '    Secure.UserQualG8F = b
   '
   '    Secure.UserSale = b
   '    Secure.UserSaleG1 = b
   '    Secure.UserSaleG1E = b
   '    Secure.UserSaleG1V = b
   '    Secure.UserSaleG1F = b
   '
   '    Secure.UserSaleG2 = b
   '    Secure.UserSaleG2E = b
   '    Secure.UserSaleG2V = b
   '    Secure.UserSaleG2F = b
   '
   '    Secure.UserSaleG3 = b
   '    Secure.UserSaleG3E = b
   '    Secure.UserSaleG3V = b
   '    Secure.UserSaleG3F = b
   '
   '    Secure.UserSaleG4 = b
   '    Secure.UserSaleG4E = b
   '    Secure.UserSaleG4V = b
   '    Secure.UserSaleG4F = b
   '
   '    Secure.UserSaleG5 = b
   '    Secure.UserSaleG5E = b
   '    Secure.UserSaleG5V = b
   '    Secure.UserSaleG5F = b
   '
   '    Secure.UserSaleG6 = b
   '    Secure.UserSaleG6E = b
   '    Secure.UserSaleG6V = b
   '    Secure.UserSaleG6F = b
   '
   '    Secure.UserSaleG7 = b
   '    Secure.UserSaleG7E = b
   '    Secure.UserSaleG7V = b
   '    Secure.UserSaleG7F = b
   '
   '    Secure.UserSaleG8 = b
   '    Secure.UserSaleG8E = b
   '    Secure.UserSaleG8V = b
   '    Secure.UserSaleG8F = b
   Put #iFreeIdx, iList, SecPw
   If Err = 0 Then
      Put #iFreeDbf, iList, Secure
      SysMsg "The New User Has Been Added.", True
      bUserInstalled = True
      iUserIdx = iList
   Else
      MsgBox "Could Not Install The New User.", _
         vbExclamation, Caption
      bUserInstalled = False
   End If
   iCurrentRec = iOldrec
   Unload Me
   
End Sub
