VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EsiAdmn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator Password"
   ClientHeight    =   4035
   ClientLeft      =   2790
   ClientTop       =   3510
   ClientWidth     =   5325
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   3720
   End
   Begin VB.TextBox txtVer 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Verify SQL Server Administrator's Password (15 Char max)"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Requires a Key to open.  Call ESI."
      Top             =   2280
      Width           =   1215
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4920
      Top             =   3240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4035
      FormDesignWidth =   5325
   End
   Begin VB.TextBox txtLog 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "SQL Server Administrator"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Update for this user only"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtPsw 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "SQL Server Administrator's Password (15 Char max)"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4320
      TabIndex        =   4
      Top             =   0
      Width           =   875
   End
   Begin VB.Label lblAdmn 
      BackStyle       =   0  'Transparent
      Caption         =   $"EsiAdmn.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lblAdmn 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Password"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Function Pass Key"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Administrator Logon"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Administrator Password"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblAdmn 
      BackStyle       =   0  'Transparent
      Caption         =   "This Function Does Not Change The SQL Server Administrator Password.  It Is To Be Used For Logon Purposes Only."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "EsiAdmn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'4/16/99 to allow a secure password
'3/5/02  change the opening labels and added a verification
'PassKey 550-041-330 (Same as security)

Dim bCancel As Byte
Dim sKey As String
Dim sAdminName As String
Dim sAdminPassword As String

Private Sub cmdCan_Click()
   bCancel = 1
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


'refined 3/5/02

Private Sub cmdUpd_Click()
   Dim A As Integer
   Dim I As Integer
   Dim bResponse As Byte
   Dim sNewPassword As String
   Dim sMsg As String
   
   bResponse = VerifyPassword()
   If bResponse = 1 Then
      If Len(Trim(txtLog)) = 0 Then txtLog = "sa"
      sMsg = "This Function Establishes The Logon For This User." & vbCr _
             & "If The Logon Is Not Setup Properly In SQL Server " & vbCr _
             & "Then The User Will Not Be Able To Perform Normally. " & vbCr _
             & "Do You Really Wish To Continue?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, sSysCaption)
      If bResponse = vbYes Then
         sNewPassword = Trim(txtPsw)
         'SaveSetting "UserObjects", "System", "NoReg", Trim(txtLog)
         SaveUserSetting USERSETTING_SqlLogin, Trim(txtLog)
         PutDatabasePassword sNewPassword
         MsgBox "System Logons Recorded.", vbInformation, sSysCaption
         Unload Me
      End If
   Else
      MsgBox "Passwords Do Not Match.", vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   'sAdminName = GetSetting("UserObjects", "System", "NoReg", sAdminName)
   sAdminName = GetUserSetting(USERSETTING_SqlLogin)
   If Trim(sAdminName) = "" Then sAdminName = "sa"
   sKey = "550-041-330"
   txtKey.PasswordChar = Chr$(176)
   txtPsw.PasswordChar = Chr$(176)
   txtVer.PasswordChar = Chr$(176)
   txtLog = sAdminName
   txtKey = String(11, " ")
   txtPsw = String(14, " ")
   txtVer = String(14, " ")
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   EsiLogon.optAdmn.value = vbUnchecked
   EsiLogon.txtPas.SetFocus
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set EsiAdmn = Nothing
   
End Sub



Private Sub Timer1_Timer()
   Static b As Byte
   b = b + 1
   If b < 10 Then
      If lblAdmn(1).Visible Then lblAdmn(1).Visible = False _
                 Else lblAdmn(1).Visible = True
   Else
      lblAdmn(1).Visible = True
      Timer1.Enabled = False
   End If
   
End Sub

Private Sub txtKey_GotFocus()
   SelectFormat Me
   
End Sub


Private Sub txtKey_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtKey_LostFocus()
   txtKey = Trim(txtKey)
   If bCancel = 0 Then CheckKey
   
End Sub


Private Sub txtLog_Change()
   If Len(txtLog) > 30 Then Beep
   
End Sub

Private Sub txtLog_GotFocus()
   SelectFormat Me
   If Len(txtPsw) = 0 Then txtPsw = "***************"
   
End Sub


Private Sub txtLog_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtLog_LostFocus()
   txtLog = Trim(txtLog)
   If Len(txtLog) > 30 Then txtLog = Left(txtLog, 30)
   
End Sub


Private Sub txtPsw_Change()
   If Len(txtPsw) > 15 Then Beep
   
End Sub


Private Sub txtPsw_GotFocus()
   If Trim(txtPsw) = "" Then txtKey = " "
   SelectFormat Me
   bInsertOn = False
   
End Sub

Private Sub txtPsw_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtPsw_LostFocus()
   txtPsw = Compress(txtPsw)
   If Len(txtPsw) > 15 Then txtPsw = Left(txtPsw, 15)
   
End Sub



Private Sub CheckKey()
   ReSize1.Enabled = False
   If txtKey <> sKey Then
      MsgBox "Invalid Pass. Call ESI.", vbExclamation, sSysCaption
   Else
      txtLog.Enabled = True
      txtPsw.Enabled = True
      txtVer.Enabled = True
      cmdUpd.Enabled = True
      On Error Resume Next
      txtLog.SetFocus
      txtKey.Enabled = False
   End If
   
End Sub

Private Function VerifyPassword() As Byte
   If Trim(txtPsw) = Trim(txtVer) Then
      VerifyPassword = 1
   Else
      VerifyPassword = 0
   End If
   
End Function

Private Sub txtVer_Change()
   If Len(txtPsw) > 15 Then Beep
   
End Sub


Private Sub txtVer_GotFocus()
   If Trim(txtVer) = "" Then txtVer = " "
   SelectFormat Me
   bInsertOn = False
   
End Sub


Private Sub txtVer_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtVer_LostFocus()
   txtVer = Compress(txtVer)
   If Len(txtVer) > 15 Then txtVer = Left(txtPsw, 15)
   bSQLOpen = 0
   On Error Resume Next
'   RdoCon.Close
   EsiLogon.lblSqlServer = "SQL Server Is Not Connected"
   EsiLogon.lblSqlServer.Refresh
   
End Sub
