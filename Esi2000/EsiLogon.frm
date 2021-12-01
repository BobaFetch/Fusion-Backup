VERSION 5.00
Begin VB.Form EsiLogon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log On To Fusion ERP"
   ClientHeight    =   4290
   ClientLeft      =   2790
   ClientTop       =   2445
   ClientWidth     =   4545
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   HelpContextID   =   40
   Icon            =   "EsiLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox OptDdb 
      Caption         =   "Save As Default Database"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "For This User Only"
      Top             =   3540
      Width           =   2775
   End
   Begin VB.CheckBox optAdmn 
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbDbs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Available DataBases (Caution: Includes All But System Databases)"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Cancel Log On"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Log On"
      Default         =   -1  'True
      Height          =   360
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Log On To Fusion ERP"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox optSve 
      Caption         =   "Save User ID"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Saves Current (Valid) Logon Information"
      Top             =   2820
      Width           =   3015
   End
   Begin VB.TextBox txtUsr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Case Insensitive Max (30) Char"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "."
      TabIndex        =   2
      ToolTipText     =   "Case Sensitive Max (15) Char"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Timer tmr3 
      Interval        =   60000
      Left            =   2700
      Top             =   4620
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version x.y.z 11/22/2007"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   1245
      Width           =   4515
   End
   Begin VB.Label lblSqlServer 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "look down here V"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Available DataBases (Caution: Includes All But System Databases)"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   0
      Picture         =   "EsiLogon.frx":014A
      ToolTipText     =   "ES/2000 ERP "
      Top             =   0
      Width           =   4530
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   1335
   End
   Begin VB.Label cmdUsr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Settings        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "User Setup And Preferences (Local Settings)"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label cmdDsn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DSN         "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "DataSource Name Setup (For Reports)"
      Top             =   4680
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "EsiLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of            ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'All new 2/28/02
'Note: Image fonts are Are Arial Black,Bold Italic and 12
'Images are hidden below form and not visible /bitmaps/mom2005,mom2006
Dim bOnLoad As Byte
Dim bKeyTab As Byte
Dim bGoodPassword As Byte

Dim iUserTries As Integer
Dim iTimer As Integer
Dim sLastPw As String
Dim slastUser As String


Private Const SHARD_PATH = &H2&

Private Declare Function SHAddToRecentDocs Lib "shell32" (ByVal dwFlags As Long, _
   ByVal dwData As String) As Long

Private Sub cmbDbs_Click()
   bSQLOpen = 0
   
End Sub

Private Sub cmbDbs_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub

Private Sub cmbDbs_LostFocus()
   cmbDbs = Trim(cmbDbs)
   If Len(cmbDbs) = 0 Then cmbDbs = cmbDbs.List(0)
   SaveDbInRegistry
'   SaveUserSetting USERSETTING_DatabaseName, cmbDbs
'   SaveUserSetting USERSETTING_SqlDsn, cmbDbs
'   sDataBase = cmbDbs
   bSQLOpen = 0
   
End Sub


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Do You Really Want To Quit " & sSysCaption & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, sSysCaption)
   If bResponse = vbYes Then
      bGoodPassword = 0
      bResponse = CloseManager()
   End If
   
End Sub

Private Sub cmdCan_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cmdCan_Click
   
End Sub


Private Sub cmdCan_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub

'Removed 3/11/02 Building out own

Private Sub cmdDsn_Click()
   On Error Resume Next
   'diaDsn.Show
   
End Sub


'Reset the ToolBars to VbNormal and inserted a Hide
'Blocked the WindowState in this routine
'For Jevco 2/12/02

Private Sub cmdOk_Click()
   'On Error Resume Next
   
   Dim shFlag As Long
   Dim shData As String
   shFlag = SHARD_PATH
   shData = App.Path & "\esi2000.exe"
   Call SHAddToRecentDocs(shFlag, shData)
   LogOnToEs2000
   
   SaveDbInRegistry

End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cmdOk_Click
End Sub

Private Sub cmdOk_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii, Me
End Sub

Private Sub cmdUsr_Click()
   On Error Resume Next
   EsiSetup.Show
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   Err.Clear
   On Error Resume Next
   'On Error GoTo modErr1

   If bOnLoad = 1 Then
      MouseCursor 13
      b = CheckSecuritySettings(True)
      'sserver = UCase$(GetSetting("Esi2000", "System", "ServerId", sserver))
      ' DNS sserver = UCase(GetUserSetting(USERSETTING_ServerName))
      sserver = UCase(GetConfUserSetting(USERSETTING_ServerName))
      
      If sserver = "" Then
         MsgBox "Server Id Isn't Set." _
            & vbCr _
            & "Select User and Then Your Server.", vbInformation, sSysCaption
         On Error Resume Next
         EsiSetup.Show
         EsiSetup.cmbSrv.SetFocus
      Else
         'If Len(Trim(txtUsr)) <> 0 Then cmdOk.SetFocus
         txtPas.SetFocus
      End If
      bOnLoad = 0
      
      'get exe creation date
      Dim vDate, sDate As String
      
      vDate = FileDateTime(App.Path & "\" & App.ExeName & ".exe")
      If vDate <> "" Then
         sDate = " " & Format(vDate, "mm/dd/yyyy")
      End If
      
      lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision & sDate
      
      optSve.BackColor = Me.BackColor
      OptDdb.BackColor = Me.BackColor
      
   End If
   
   If txtUsr.Enabled Then txtUsr.SetFocus
   If optAdmn.Value = vbChecked Then
      optAdmn.Value = vbUnchecked
      Unload EsiAdmn
   End If

'   Exit Sub
'
'modErr1:
'   MsgBox ("LoginError: " & Err.Description)
'   MsgBox ("LoginError: " & Err.Number)
'   Exit Sub
   
End Sub

Private Sub Form_Initialize()
   'All new 2/28/02
   'Note: Image fonts are Are Arial Black,Bold Italic and 12
   'Images are hidden below form and not visible /bitmaps/mom2005,mom2006
   'SetFormSize Me
   '    Dim sYear As String
   '   'change banner for year
   '    sYear = Format$(Now, "yyyy")
   '    Select Case sYear
   '        Case "2005"
   '            Image1(0).Picture = Image1(3).Picture
   '        Case "2006"
   '            Image1(0).Picture = Image1(4).Picture
   '        Case "2007"
   '            Image1(0).Picture = Image1(0).Picture
   '        Case Else
   '            Image1(0).Picture = Image1(3).Picture
   '    End Select
   '''txtPas.PasswordChar = Chr$(176)
   'Caption = "Log On To ES/2000 ERP "
   Me.BackColor = RGB(212, 208, 200)
   cmdOk.BackColor = RGB(212, 208, 200)
   cmdCan.BackColor = RGB(212, 208, 200)
   txtPas.PasswordChar = "*"
   
   
End Sub


'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Const vbCtrlMask = 2
'    If (Shift = vbCtrlMask) And (KeyCode = 82) Then
'
'
'
'        ESIRegister.Show    'Show the registration form
'    End If
'End Sub


Private Sub Form_Load()
   Dim aRow As Byte
   Dim b As Byte
   Dim sPw As String
   Dim sDefDb As String
   
   Dim sSyslog As String
   Dim sSysPW As String
   Dim bTestDBPresent As Boolean
   Dim strDatabs As String
   
   
   Dim sData(12) As String
   
   Caption = GetSystemCaption()
   bGoodPassword = 0
   bWeAreLoaded = 0
   cmdOk.ToolTipText = "Log On To Fusion ERP"
   
   optSve.Value = GetUserSetting(USERSETTING_SaveUserLogin)
   OptDdb.Value = GetUserSetting(USERSETTING_SaveDatabase)
   slastUser = GetUserSetting(USERSETTING_UserLogin)
   
   If slastUser = "" Then slastUser = GetNetId()
   
   If optSve = vbChecked Then
      txtUsr.Enabled = False
   End If
   SaveSetting "Esi2000", "System", "Database", "Esi2000Db"
   If Len(Trim(slastUser)) = 0 Then slastUser = " "
   txtUsr = slastUser
   txtPas = ""
   bOnLoad = 1
   
   
   'Registered Databases?
   ' DNS Change
   '  bTestDBPresent = False
   '  For b = 0 To 11
   '     sData(b) = Trim(GetSetting("Esi2000", "System", "UserDatabase" & Trim(Str(b)), sData(b)))
   '     If Trim(sData(b)) <> "" Then cmbDbs.AddItem sData(b)
   '     If UCase(Trim(sData(b))) = "TESTDB" Then bTestDBPresent = True
   '  Next
   
   strDatabs = GetConfUserSetting(USERSETTING_SqlDsn)
   If Trim(strDatabs) <> "" Then cmbDbs.AddItem strDatabs
   If UCase(strDatabs) = "TESTDB" Then bTestDBPresent = True
   
   If cmbDbs.ListCount = 0 Then
      cmbDbs.AddItem "Esi2000Db"
   Else
      cmbDbs.Enabled = True
      z1(3).Enabled = True
      OptDdb.Enabled = True
   End If
   On Error Resume Next
   
'  DNS change
'   For aRow = 0 To cmbDbs.ListCount - 1
'      If Trim(cmbDbs.List(aRow)) = "Esi2000Db" Then
'         b = 1
'         Exit For
'      Else
'         b = 0
'      End If
'   Next
'   If b = 0 Then cmbDbs.AddItem "Esi2000Db"
'   cmbDbs = cmbDbs.List(0)
   
   If InTestMode Then
        If Not bTestDBPresent Then cmbDbs.AddItem "TestDB"
        cmbDbs = "TestDB"
   End If
   
   
   ' DNS sDefDb = Trim(GetUserSetting(USERSETTING_DefaultDatabase))
   sDefDb = Trim(GetConfUserSetting(USERSETTING_DefaultDatabase))
   
   If Not InTestMode And sDefDb <> "" Then cmbDbs = sDefDb
   SaveUserDatabase
   SetFocus
   
End Sub

Private Sub Form_LostFocus()
   If bGoodPassword Then Unload Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If sLastPw = Trim(txtPas) Then bGoodPassword = True
   If bGoodPassword Then
      sLastPw = Trim(txtPas)
      'SaveSetting "Esi2000", "System", "Lastuser", txtUsr
      SaveUserSetting USERSETTING_UserLogin, txtUsr
      'SaveSetting "Esi2000", "System" & "-" & Replace(App.Path, "\", "+"), "Lastuser", txtUsr
      If optSve.Value = vbChecked Then
         'SaveSetting "Esi2000", "System", "SaveId", "1"
         SaveUserSetting USERSETTING_SaveUserLogin, "1"
         sLastPw = ScramblePw(sLastPw)
         'SaveSetting "Esi2000", "System", "LastId", sLastPw
'''         SaveUserSetting USERSETTING_UserPassword, sLastPw
      Else
         'SaveSetting "Esi2000", "System", "LastId", ""
'''         SaveUserSetting USERSETTING_UserPassword, ""
         'SaveSetting "Esi2000", "System", "SaveId", "0"
         SaveUserSetting USERSETTING_SaveUserLogin, "0"
      End If
   Else
      'SaveSetting "Esi2000", "System", "LastId", ""
'''      SaveUserSetting USERSETTING_UserPassword, ""
      'SaveSetting "Esi2000", "System", "SaveId", "0"
      SaveUserSetting USERSETTING_SaveUserLogin, "0"
      bUserLoggedOn = 0
'MsgBox "0-1"
   End If
   'SaveSetting "Esi2000", "System", "SaveUserDatabase", OptDdb.value
   SaveUserSetting USERSETTING_SaveDatabase, OptDdb.Value
   'SaveSetting "Esi2000", "System", "DefaultUserDatabase", Trim(cmbDbs)
   SaveUserSetting USERSETTING_DefaultDatabase, Trim(cmbDbs)
   tmr3.Enabled = False
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   MouseCursor 0
   Set EsiLogon = Nothing
   
End Sub




Private Sub Image1_Click()
   'Note: Image fonts are Are SerpentineDBol Bold, Italic, and 12
   'Images are hidden below form and not visible /bitmaps/mom2005,mom2006
   OpenWebPage "http://www.fusionerp.net/"
   
   
End Sub


Private Sub optSve_Click()
   'On Error Resume Next
   If optSve.Value = vbChecked Then
        txtUsr.Enabled = False
      '  txtPas.Enabled = False
   Else
      txtUsr.Enabled = True
      txtUsr.SetFocus
      '''txtPas.Enabled = True
   End If
   
End Sub

Private Sub Tmr3_Timer()
   iTimer = iTimer + 1
   
End Sub

Private Sub txtPas_Change()
   If Len(txtPas) = 0 Then SelectFormat Me, True
   
End Sub

Private Sub txtPas_GotFocus()
   bKeyTab = 1
   SelectFormat Me, True
   
End Sub

Private Sub txtPas_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then bKeyTab = 0 Else bKeyTab = 1
End Sub

Private Sub txtPas_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub

Private Sub txtPas_LostFocus()
   Dim b As Byte
   txtPas = Compress(txtPas)
   '''If Trim(txtPas) = "" Then txtPas = String$(15, " ")
   If Len(Trim(txtPas)) = 0 Then bGoodPassword = 0
   If bKeyTab = 0 Then LogOnToEs2000
   
End Sub

Private Sub txtPas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bKeyTab = 1
   
End Sub

Private Sub txtUsr_Change()
   If Len(txtUsr) = 0 Then SelectFormat Me
   
End Sub

Private Sub txtUsr_GotFocus()
   bKeyTab = 1
   '''If Len(Trim(txtPas)) = 0 Then txtPas = String$(15, " ")
   If optSve.Value = vbUnchecked Then SelectFormat Me
   
End Sub

Private Sub txtUsr_KeyDown(KeyCode As Integer, Shift As Integer)
   If UCase$(txtUsr) = "ADMINISTRATOR" Then
      optAdmn.Value = vbChecked
      If KeyCode = vbKeyF8 Then EsiAdmn.Show
   End If
   
End Sub


Private Sub txtUsr_KeyPress(KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub txtUsr_LostFocus()
   txtUsr = Trim(txtUsr)
   If Len(txtUsr) > 20 Then txtUsr = Left(txtUsr, 30)
   
End Sub




Private Sub SaveUserDatabase()
   If Len(cmbDbs) = 0 Then cmbDbs = cmbDbs.List(0)
   'SaveSetting "Esi2000", "System", "CurDatabase", cmbDbs
   SaveUserSetting USERSETTING_DatabaseName, cmbDbs
End Sub

Private Sub LogOnToEs2000()
    Dim intResponse As Integer
    Dim bRegOk As Boolean
    Dim sMsg As String
    
'   'don't continue if test app against live db or vice versa
'   If Not OpenSqlServer Then
'      cmdOk.Enabled = True
'      Exit Sub
'   End If
'
   'Temp to allow demo to fire
   Dim lClose As Long
   lClose = FindWindow(vbNullString, "ESI CloseSections")
   If lClose > 0 Then SendMessage lClose, WM_CLOSE, 0&, 0&
   
   PopMsg.msg = "Logging On To Fusion ERP."
   PopMsg.Show
   EsiLogon.lblSqlServer = "Logging On To Fusion ERP"
   EsiLogon.lblSqlServer.Refresh
   cmdOk.Enabled = False
   Dim b As Byte
   Dim bGo As Boolean
   If Len(txtPas) > 0 Then bGoodPassword = 1
   sUserId = txtUsr
   sPassword = Trim(txtPas)
   MouseCursor 13
   If bSecSet = 0 Then
      
      MsgBox "Old security model no longer supported.  Contact ESI"
      bUserLoggedOn = 0
      MsgBox "No Permissions Set. Logging Off.", vbExclamation, sSysCaption
      Unload PopMsg
      Unload Me
      Exit Sub
      
   Else
      b = GetUser(Trim(txtUsr), sPassword)
      If b = 0 Then
         MouseCursor 0
         MsgBox "User Name Or Password Wasn't Found.", _
            vbExclamation, sSysCaption
         iUserTries = iUserTries + 1
         If iUserTries = 4 Then
            bUserLoggedOn = 0
            MsgBox "No Permissions Set. Logging Off.", vbExclamation, sSysCaption
            bGo = False
         End
         Exit Sub
      Else
         On Error Resume Next
         cmdOk.Enabled = True
         txtUsr.SetFocus
      End If
   Else
      bGo = True
   End If
End If
If bGo Then
   If Not OpenDBServer Then
      MouseCursor ccDefault
      cmdOk.Enabled = True
      Exit Sub
   End If
   
   UpdateTables
   ' Make connection to the database server with the ADO object.
   'OpenDBServer
   UpdateDatabase
   
' BBS Temporarity remarked out this code so that we
' can implement it when we are ready.
    'BBS Added for new registration logic (ticket #38600
'    If Not (InStr(1, UCase(Command), "FUSIONROCKS") > 0) Then
'        bRegOk = RegistrationOk(sMsg, False)
'        If Len(sMsg) > 0 Then
'            intResponse = MsgBox(sMsg, vbYesNo, "Fusion Registration")
'            If intResponse <> vbYes And bRegOk = False Then
'                CloseManager
'                Exit Sub
'            End If
'            If intResponse = vbYes Then
'                Me.Hide
'                Unload PopMsg
'                ESIRegister.Show vbModal
'                If Not RegistrationOk(sMsg) Then
'                    CloseManager
'                    Exit Sub
'                End If  ' if not regisrationok(smsg)
'            End If ' if intresponse=vbyes
'        End If ' if len(smsg)>0
'    End If  ' if not runninginide
    
   SaveUserDatabase
   sUserId = UCase$(sUserId)
   sUserId = Compress(sUserId)
   SaveSetting "Esi2000", "sections", "EsiOpen", 1
   SaveSetting "Esi2000", "system", "UserId", sUserId
   If Trim(cmbDbs) <> "" Then
      'SaveSetting "Esi2000", "System", "CurDatabase", cmbDbs
      SaveUserSetting USERSETTING_DatabaseName, cmbDbs
      'SaveSetting "Esi2000", "System", "SqlDsn", cmbDbs
      SaveUserSetting USERSETTING_SqlDsn, cmbDbs
   End If
   cmdCan.Enabled = False
   bShowVertical = GetSetting("Esi2000", "mngr", "ShowVertical", bShowVertical)
   bVerticalLoaded = bShowVertical
   bWeAreLoaded = 1
   bGoodPassword = True
   'If bVerticalLoaded = 1 Then Esi2000v.Show _
   '                     Else Esi2000h.Show
                        
   If bVerticalLoaded = 1 Then
      Esi2000v.Show
   Else
      Dim bret As Boolean
      
      bret = EnableDataColMod()
      If (bret = True) Then
         Esi2000ha.Show
      Else
         Esi2000h.Show
      End If
   End If
   
   DoEvents
   Sleep 2000
   WriteUserLog
   MouseCursor 0
   Unload Me
End If

End Sub

Private Sub txtUsr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bKeyTab = 1
   
End Sub

Private Sub SaveDbInRegistry()
   sDataBase = cmbDbs
   SaveUserSetting USERSETTING_DatabaseName, sDataBase
   SaveUserSetting USERSETTING_SqlDsn, sDataBase
End Sub
