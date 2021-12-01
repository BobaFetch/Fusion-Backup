Attribute VB_Name = "Esimgr"
'*** ES/2000 (ES/2001 - ES/2007)) is the property of          ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit

'Global system calls
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'SendMessage:
Public Const WM_CLOSE = &H10

'ini stuff
'Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'7.1.5 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
   ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_TAB = &H9

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long


'Mail, Web, Apps etc
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const Hwnd_NOTOPMOST = -2
Public Const HWND_TOP = 0
'Public Const hWnd_TopMost = -1
Public Const Swp_NOSIZE = &H1
Public Const Swp_NOMOVE = &H2

'ESI Color Constants
'Public Const ES_RED = &HC0                 'jse red for labels
'Public Const ES_BLUE = &H800000            'jse blue for labels
'Public Const Es_FormBackColor = &H8000000F
'Public Const Es_CheckBoxForeColor = &H8000000F
'Public Const Es_TextBackColor = &H80000005
'Public Const Es_TextForeColor = &H80000008
'Public Const Es_CustomBackColor = &H80000001
'Public Const ES_ViewBackColor = &HE0FFFF '11/17/04 RGB(255, 253, 223)
'other
'Public Const ES_NOQUESTION = &H124        'Question and return
'Public Const ES_YESQUESTION = &H24  'Question and return (Default YES)

'About box
'Public Type MEMORYSTATUS
'        dwLength As Long
'        dwMemoryLoad As Long
'        dwTotalPhys As Long
'        dwAvailPhys As Long
'        dwTotalPageFile As Long
'        dwAvailPageFile As Long
'        dwTotalVirtual As Long
'        dwAvailVirtual As Long
'End Type

'Decimal Conversions
Private RdoTest As ADODB.Recordset
Private RdoCol As ADODB.Recordset
'Private ER As rdoError

Private bByte As Byte
Private bBadCol As Byte
Private iColCounter As Integer
Private sConstraint As String

Public bUpdated As Byte
Dim sBadColumns(700) As String

'Menu Constants
Public Const SC_CLOSE = &HF060
Public Const MF_GRAYED = &H1
'Public Const SC_MAXIMIZE = &HF030
'Public Const SC_MINIMIZE = &HF020
'Public Const SC_MOVE = &HF010
'Public Const GWL_STYLE = (-16)
'Public Const WS_MINIMIZEBOX = &H20000

'Cursor types
Public Const ES_FORWARD = 0 'Default
Public Const ES_KEYSET = 1
Public Const ES_DYNAMIC = 2
Public Const ES_STATIC = 3

'Public RdoCon As rdoConnection
'Public RdoEnv As rdoEnvironment
Public clsADOCon As ClassFusionADO

Public bGoodPassword As Boolean
Public bEnterAsTab As Byte
Public bInsertOn As Byte
Public bOpenSection As Byte
Public bNoResize As Byte
Public bShowVertical As Byte
Public bSqlRows As Byte
Public bUserLoggedOn As Byte
Public bVerticalLoaded As Byte
Public bWeAreLoaded As Byte
Public bSQLOpen As Byte
Public iMinimize As Integer
Public iSaveCurrent As Integer

Public lLogId As Long

'Public sDataBase    As String
'Public sDsn         As String
Public sPassword As String
Public sFilePath As String
Public sHelpType As String
Public sInitials As String
Public sReportPath As String
'Public sSysCaption  As String
'Public sserver      As String
Public sSql As String
Public sUserId As String
Public sUserName As String
Public sWorkStation As String
Public ES_SYSDATE As Variant 'Server Date/Time

Public Type SectionInfoType
   AppTitle As String
   AppName As String
   ExeName As String
   SectionOpen As Double
   AppDescription As String
   AppButtons As String 'separated by vbcr's
End Type

Public SectionInfo(1 To 10) As SectionInfoType

Type tUser
   Adduser As Integer
   Level As Integer
   Group1 As Integer
   Group2 As Integer
   Group3 As Integer
   Group4 As Integer
   Group5 As Integer
   Group6 As Integer
   Group7 As Integer
   Group8 As Integer
End Type

Public User As tUser

''8/26/06 Compliance with Sections
'Type CurrentSelections
'    CurrentPart     As String
'    CurrentVendor   As String
'    CurrentCustomer As String
'    CurrentShop     As String
'    CurrentRegion   As String
'    CurrentGroup    As String
'    CurrentUser     As String
'End Type
'Public cUR As CurrentSelections


'For column updates

Private Function GetConstraint(sDescription As String) As String
'   GetConstraint = ""
'   bByte = InStr(1, sDescription, "DF_")
'   If bByte > 0 Then
'      GetConstraint = Mid$(sDescription, bByte, Len(sDescription))
'      bByte = InStr(1, GetConstraint, "'")
'      If bByte > 0 Then GetConstraint = Left$(GetConstraint, bByte - 1)
'   Else
'      bByte = InStr(1, sDescription, "DEFZERO")
'      If bByte > 0 Then
'         sSql = "sp_unbindefault '" & RdoCol.rdoColumns(2) & "." & RdoCol.rdoColumns(3) & "'"
'         RdoCon.Execute sSql, rdExecDirect
'      End If
'   End If
   
   
End Function


'2/9/04 If Trim(txtEml) <> "" Then OpenWebPage Trim(txtWeb)
'SendTo = request and doesn't have to be a web address

Public Sub OpenWebPage(SendTo As String)
   Dim lTask As Long
   On Error Resume Next
   lTask = ShellExecute(EsiLogon.hwnd, "open", SendTo, _
           vbNullString, vbNullString, SW_SHOW)
   If lTask < 32 Then MsgBox SendTo & " doesn't exist or couldn't be opened." & vbCr _
      & "If the file is a Log, then there may be no data to report.", _
      vbInformation, "ES/2000 ERP"
   
   '//ShellExecute(
   'hwnd>> handle to parent window
   'lpOperation>> pointer to string that specifies operation to perform
   'lpFile>> pointer to filename or folder name string
   'lpParameters>> pointer to string that specifies executable-file parameters
   'lpDirectory>> pointer to string that specifies default directory
   ' nShowCmd>>whether file is shown when opened
   
End Sub

Public Function CheckLen(sTextBox As String, iTextLength As Integer) As String
   sTextBox = Trim(sTextBox)
   If Len(sTextBox) > iTextLength Then
      'SysSysbeep
      sTextBox = Left(sTextBox, iTextLength)
   End If
   CheckLen = sTextBox
   
End Function

Public Function EnableDataColMod() As Boolean
   
   Dim RdoMod As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT ISNULL(CODATACOLMOD, 0) FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMod, ES_FORWARD)
   If bSqlRows Then
      EnableDataColMod = IIf(RdoMod.Fields(0) = 1, True, False)
   Else
      EnableDataColMod = False
   End If
   
   Set RdoMod = Nothing
   Exit Function
   
modErr1:
   EnableDataColMod = False

End Function


Public Function StrCase(sTextStr As Variant, Optional bTextOption As Byte)
   Dim bAutoCaps As Byte
   Dim iKeyState As Integer
   Dim iStrLen As Integer
   Dim sNewStr As String
   If bAutoCaps = 1 Or iKeyState = 1 Then
      StrCase = sTextStr
   Else
      sNewStr = Trim$(sTextStr)
      iStrLen = Len(sNewStr)
      If iStrLen > 1 Then
         If Asc(Mid(sNewStr, 2, 1)) > 64 And Asc(Mid(sNewStr, 2)) < 90 Then
            StrCase = sNewStr
            Exit Function
         End If
      End If
      If iStrLen > 0 Then sNewStr = StrConv(sNewStr, vbProperCase)
      StrCase = sNewStr
   End If
   
End Function

'2/26/02 find the TestDb (if any and add to user list)
'Use these registery settings
'Get this:
'sDataBase = GetSetting("Esi2000", "System", "CurDatabase", sDataBase)
'Save these
'SaveSetting "Esi2000", "System", "UserDatabase1", sDataBases(0)
'SaveSetting "Esi2000", "System", "UserDatabase2", sDataBases(1)
'....to 12
'1/14/04 expanded use if TestDb to any but sys_databases

'Public Sub GetDataBases()
'   Dim RdoDbs As rdoResultset
'   Dim aRow As Byte
'   Dim b As Byte
'   Dim sDataBases(12) As String
'   Dim sData As String
'
'   MouseCursor 13
'   On Error Resume Next
'   sSql = "sp_databases"
'   bSqlRows = GetDataSet(RdoDbs, ES_FORWARD)
'   b = 0
'   If bSqlRows Then
'      With RdoDbs
'         Do Until .EOF
'            If Not IsNull(.rdoColumns(0)) Then
'               sData = UCase$(Trim(.rdoColumns(0)))
'               If sData <> "MASTER" Then
'                  If sData <> "NORTHWIND" Then
'                     If sData <> "PUBS" Then
'                        If sData <> "TEMPDB" Then
'                           If sData <> "MODEL" Then
'                              If sData <> "MSDB" Then
'                                 Err.Clear
'                                 sSql = "USE " & sData
'                                 RdoCon.Execute sSql, rdExecDirect
'
'                                 sSql = "Qry_FillRegions"
'                                 RdoCon.Execute sSql, rdExecDirect
'                                 If Err = 0 Then
'                                    b = b + 1
'                                    sDataBases(b) = "" & Trim(.rdoColumns(0))
'                                 End If
'                                 If b > 11 Then Exit Do
'                              End If
'                           End If
'                        End If
'                     End If
'                  End If
'               End If
'            End If
'            .MoveNext
'         Loop
'         .Cancel
'      End With
'   End If
'   'Save these settings for next future
'   For aRow = 0 To 11
'      SaveSetting "Esi2000", "System", "UserDatabase" & Trim(Str(aRow)), sDataBases(aRow)
'      If Trim(sDataBases(aRow)) <> "" Then RegisterTestDbDsn sDataBases(aRow)
'   Next
'   Set RdoDbs = Nothing
'   sSql = "USE " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
'
'End Sub

Public Function GetServerDateTime() As Variant
   Dim RdoTme As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT GETDATE()"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTme, ES_FORWARD)
   If bSqlRows Then GetServerDateTime = RdoTme.Fields(0)
   Set RdoTme = Nothing
   Exit Function
   
modErr1:
   GetServerDateTime = Format(Now, "mm/dd/yy")
   
End Function

'Standard procedure for receiving resultsets
'bSqlRows = GetDataSet (RdoRes, ES_FORWARD)
'See also GetQuerySet for VB built queries

'Public Function GetDataSet(RdoDataSet As rdoResultset, Optional iCursorType As Integer) As Byte
'   ' Use local error Trapping
'   'RdoCon.QueryTimeout = 40
'   Err = 0
'   If iCursorType = ES_FORWARD Then
'      'Forward only "cursor" (not a cursor)
'      Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
'      If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'         GetDataSet = 1
'      Else
'         GetDataSet = 0
'      End If
'   Else
'      If iCursorType = ES_KEYSET Then
'         'Keyset cursor for Editing
'         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = 1
'         Else
'            GetDataSet = False
'         End If
'      ElseIf iCursorType = ES_DYNAMIC Then
'         'Dynamic
'         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenDynamic, rdConcurRowVer)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = 1
'         Else
'            GetDataSet = 0
'         End If
'      ElseIf iCursorType = ES_STATIC Then
'         'Static Cursor - Note: Needed for BLOBS
'         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenStatic, rdConcurReadOnly)
'         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
'            GetDataSet = 1
'         Else
'            GetDataSet = 0
'         End If
'      End If
'   End If
'   If Err > 0 Then GetDataSet = 0
'
'
'End Function
'

'Public Function RunningInIDE() As Byte
'   'Test to see if we are in VB or user mode
'   'Check to see where the program is running.
'   'Assume that we are not running in VB5 for now
'   'Calling Debug is ignored except in VB, so will produce
'   'an error...
'   RunningInIDE = 0
'
'   'Exit Function '@@@temp test
'
'   On Error GoTo ModErr
'   Debug.Print 1 / 0
'   Exit Function
'
'ModErr:
'   On Error GoTo 0
'   'yep, it's VB
'   RunningInIDE = 1
'   'sFilePath = Mid(App.Path, 1, LastIndexOf(App.Path, "\")) + "ReleaseCandidates\"
'
'End Function

Public Function LastIndexOf(StringToSearch As String, StringToFind As String) As Integer
   'find last instance of search string
   '= 0 if no instance found
   Dim iStart As Integer, iFound As Integer
   iFound = 0
   iStart = 1
   Do While iStart > 0
      iStart = InStr(iStart, StringToSearch, StringToFind, vbTextCompare)
      If iStart > 0 Then
         iFound = iStart
         iStart = iStart + 1
      End If
   Loop
   LastIndexOf = iFound
End Function

'Does the file exist?
'Make one if not
'Only used in mom

Public Sub CheckSecurity()
   Dim bTry As Byte
   Dim a As Integer
   Dim sCoServer As String
   MouseCursor 13
   
   bRetry = 0
   '    If sServer = "DEVSERVER\DEVSERVER" Then
   '        sCoServer = "c:\esi2000\"
   '    Else
   '        sCoServer = App.Path & "\"
   '    End If
   
   sCoServer = App.Path & "\" '8/3/07 - always run in local directory
   
   On Error GoTo modErr1
   Open sCoServer & "rstval.eid" For Random Shared As #22 Len = Len(SecPw)
   a = LOF(22)
   'sserver = UCase$(GetSetting("Esi2000", "System", "ServerId", sserver))
   ' DNS sserver = UCase(GetUserSetting(USERSETTING_ServerName))
   sserver = UCase(GetConfUserSetting(USERSETTING_ServerName))
   If a = 0 And sserver <> "" Then
      'First record identifies the company
      Open sCoServer & "rstval.edd" For Random Shared As #23 Len = Len(Secure)
      SecPw.PassWord = sserver
      SecPw.UserAdmn = 0
      SecPw.UserRecord = 1
      Put #22, 1, SecPw
      Secure.UserName = sserver
      InitializePermissions Secure, 0
      
      Put #23, 1, Secure
      SetDefaultUsers 22, 23
   Else
      'What if the company file gets shipped. dated record too"
      Get #22, 1, SecPw
      If a = 399 Then
         'Need to dump it and try again
         Close #22
         Close #23
         'Changed to allow manual dumping of files 11/07/01
         SetAttr sCoServer & "rstval.eid", vbNormal
         SetAttr sCoServer & "rstval.edd", vbNormal
         
         'Kill sCoServer & "rstval.eid"
         'Kill sCoServer & "rstval.edd"
         bRetry = 1
         Exit Sub
      End If
   End If
   a = a \ Len(SecPw)
   '    If a > 3 Then
   '        bSecSet = 1
   '    Else
   '        bSecSet = 0
   '    End If
   bSecSet = 1 'always advanced security
   
   
   Close #22
   Close #23
   SetAttr sCoServer & "rstval.eid", vbHidden + vbArchive + vbSystem
   SetAttr sCoServer & "rstval.edd", vbHidden + vbArchive + vbSystem
   Exit Sub
   
modErr1:
   bTry = bTry + 1
   If bTry < 3 Then Resume Next
   
End Sub

'Install built in users

Public Sub SetDefaultUsers(File1 As Integer, File2 As Integer)
   Dim i As Integer
   Dim sDefaultPw As String
   
   'user 2 = ESIADMN
   'user 3 = SYSMGR
   
   sDefaultPw = "kihei"
   sDefaultPw = ScramblePw(sDefaultPw)
   SecPw.PassWord = sDefaultPw
   SecPw.UserLcName = "esiadmn"
   SecPw.UserUcName = "ESIADMN"
   SecPw.UserAdmn = 1
   SecPw.UserRecord = 2
   Put #File1, 2, SecPw
   
   '    sDefaultPw = "051681"
   '    sDefaultPw = ScramblePw(sDefaultPw)
   '    SecPw.PassWord = sDefaultPw
   '    SecPw.UserLcName = "larryh"
   '    SecPw.UserUcName = "LARRYH"
   '    SecPw.UserAdmn = 1
   '    SecPw.UserRecord = 3
   '    Put #File1, 3, SecPw
   '
   sDefaultPw = "sysmgr"
   sDefaultPw = ScramblePw(sDefaultPw)
   SecPw.PassWord = sDefaultPw
   SecPw.UserLcName = "sysmgr"
   SecPw.UserUcName = "SYSMGR"
   SecPw.UserAdmn = 1
   SecPw.UserRecord = 3
   Put #File1, 3, SecPw
   
   
   
   Secure.UserName = "ESI-SEG Administrator"
   Secure.UserInitials = "ESI"
   Secure.UserNickName = "Software Guy"
   Put #File2, 2, Secure
   
   Secure.UserName = "System Manager"
   Secure.UserInitials = "MGR"
   Secure.UserNickName = "SysMgr"
   Put #File2, 3, Secure
   
   For i = 2 To 3
      Get #File2, i, Secure
      Secure.UserActive = 1
      Secure.UserAddUser = 1
      Secure.UserLevel = 0
      Secure.UserNew = Format$(Now, "mm/dd/yy")
      
      InitializePermissions Secure, 1
      
      Put #File2, i, Secure
   Next
   
End Sub

'Journals Open? Means there no Accounts either.

Public Function CheckJournals() As Byte
   Dim RdoJrn As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT MJTYPE FROM JrhdTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoJrn, ES_FORWARD)
   If bSqlRows Then
      CheckJournals = 1
      RdoJrn.Cancel
   Else
      CheckJournals = 0
   End If
   Set RdoJrn = Nothing
   Exit Function
   
DiaErr1:
   CheckJournals = 0
   
End Function

' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory

Public Function GetWindowsDir()
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Integer
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   
   '
   'Get the windows directory and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetWindowsDir = strBuf
   Else
      GetWindowsDir = ""
   End If
   
End Function

'gets password incription (self documenting)
'if you can't figure it out, then you need not know

Public Function GetPassword(sPassword As String) As String
   Dim k As Integer
   Dim i As Integer
   Dim n As Integer
   Dim sNewPw As String
   
   sNewPw = ""
   On Error Resume Next
   k = Len(Trim$(sPassword))
   n = 79
   For i = 1 To k
      n = n + 1
      Mid$(sPassword, i, 1) = Chr$(Asc(Mid$(sPassword, i, 1)) - n)
   Next
   For i = k To 1 Step -1
      sNewPw = sNewPw & Mid$(sPassword, i, 1)
   Next
   GetPassword = sNewPw
   
End Function

'sets password incription (self documenting)

Public Function SetPassword(sTestPassword As String) As String
   Dim k As Integer
   Dim i As Integer
   Dim n As Integer
   Dim sNewPw As String
   On Error Resume Next
   k = Len(Trim$(sTestPassword))
   For i = k To 1 Step -1
      sNewPw = sNewPw & Mid$(sTestPassword, i, 1)
   Next
   n = 79
   For i = 1 To k
      n = n + 1
      Mid$(sNewPw, i, 1) = Chr$(Asc(Mid$(sNewPw, i, 1)) + n)
   Next
   SetPassword = sNewPw
   
End Function


Sub SysMsg(SysMessage As String, TimerOn As Byte, Optional frm As Form)
   'Timer On either True or False
   'Syntax: SysMsg "User Message", True
   'SysMessage may be up to 24 characters
   On Error GoTo SysErr1
   PopMsg.Top = EsiLogon.Top + 500
   PopMsg.tmr1.Enabled = TimerOn
   PopMsg.msg = SysMessage
   PopMsg.Show
   On Error Resume Next
   frm.Refresh
   Exit Sub
   
SysErr1:
   Resume SysErr2
SysErr2:
   On Error Resume Next
   PopMsg.tmr1.Enabled = TimerOn
   PopMsg.msg = SysMessage
   PopMsg.Show
   DoEvents
   
End Sub

Public Sub KeyCheck(KeyAscii As Integer)
   If KeyAscii = 13 Then keybd_event VK_TAB, 0, 0, 0
   
End Sub


Public Sub SelectFormat(frm As Form, Optional DontInsertBlank As Boolean)
   'Selects all of the text in fixed length text boxes
   'and combo boxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   'Allow for the possibility that the control has been
   'disabled
   'DontInsertBlank = True to skip insertion of a blank into the field
   'this is required for password entry (10/17/08)
   
   On Error Resume Next
   If Not DontInsertBlank Then
      If Len(frm.ActiveControl.Text) = 0 Then frm.ActiveControl.Text = " "
   End If
   frm.ActiveControl.SelStart = 0
   frm.ActiveControl.SelLength = Len(frm.ActiveControl)
   
End Sub

Public Function CheckOpen(iSection As Integer) As Byte
   Dim lCheckApp As Long
   CheckOpen = False
   Dim iOpen As Integer
   '    Dim sSection As String
   '    Select Case iSection
   '        Case 1
   '            sSection = "admn"
   '        Case 2
   '            sSection = "prod"
   '        Case 3
   '            sSection = "engr"
   '        Case 4
   '            sSection = "sale"
   '        Case 5
   '            sSection = "fina"
   '        Case 6
   '            sSection = "qual"
   '        Case 7
   '            sSection = "invc"
   '        Case Else
   '            sSection = ""
   '    End Select
   lCheckApp = FindWindow(vbNullString, "ESI CloseSections")
   If lCheckApp > 0 Then SendMessage lCheckApp, WM_CLOSE, 0&, 0& '*    End
   iOpen = GetSetting("Esi2000", "sections", SectionInfo(iSection).AppTitle, iOpen)
   'If iOpen = 0 Then SectionInfo(iSection) = 0
   If iOpen = 0 Then SectionInfo(iSection).SectionOpen = 0
   
End Function


'3/19/07 Blocked (not used)

Function GetIniString$(IniFile$, Section$, Entry$, Default$)
   '    Dim nChars
   '    Dim mzBuff$
   '
   '    If Len(mzBuff) < 255 Then mzBuff = Space(255)
   '    nChars = GetPrivateProfileString(Section$, Entry$, Default$, mzBuff$, 255, IniFile$)
   '    If nChars > 0 Then
   '        GetIniString$ = Left(mzBuff, nChars)
   '    Else
   '        GetIniString$ = ""
   '    End If
   
End Function

Sub MouseCursor(MCursor As Byte)
   'Allows consistant MousePointer Updates
   Screen.MousePointer = MCursor
   
End Sub

Sub KeyCase(KeyAscii As Integer, frm As Form)
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
      End If
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If Len(frm.ActiveControl) > 0 Then _
             If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
   End If
   
End Sub

Public Sub Main()
   Dim b As Byte
   Dim i As Integer
   Dim l As Long
   Dim sAppTitle As String
   Dim sEnterAsTab As String
   Dim sAppPath As String
   Dim sWindows As String
   Dim sTest As String
   
   Dim iRed As Integer
   Dim iGreen As Integer
   Dim iBlue As Integer
   
   If App.PrevInstance Then
      sAppTitle = App.Title
      App.Title = "EsiMan"
      AppActivate sAppTitle
      SendKeys "% R", True
      End
      Exit Sub
   End If

   On Error Resume Next
    bIdeRunning = IIf(RunningInIDE(), 1, 0) 'VB Running
   sWindows = GetWindowsDir() & "\" & "system\"
   sAppPath = App.Path & "\" '8/3/07 - always run in local directory

   '7/14/05
   CopyHelp
   'reset registry to mark all as closed
   SaveSetting "Esi2000", "Sections", "admn", 0
   SaveSetting "Esi2000", "Sections", "prod", 0
   SaveSetting "Esi2000", "Sections", "engr", 0
   SaveSetting "Esi2000", "Sections", "sale", 0
   SaveSetting "Esi2000", "Sections", "fina", 0
   SaveSetting "Esi2000", "Sections", "qual", 0
   SaveSetting "Esi2000", "Sections", "invc", 0
   sSysCaption = GetSystemCaption
   
   'MdiSection Color Settings (Save blanks as zero)
   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
   
   SaveSetting "Esi2000", "System", "SectionBackColorR", iRed
   SaveSetting "Esi2000", "System", "SectionBackColorG", iGreen
   SaveSetting "Esi2000", "System", "SectionBackColorB", iBlue
   
   'Use this to keep the CD-ROM and Build current
   bNoResize = GetSetting("Esi2000", "System", "ResizeForms", bNoResize)
   If bNoResize = 0 Then
      SaveSetting "Esi2000", "System", "ResizeForms", 2
      SaveSetting "Esi2000", "System", "ResizeForm", 0
   End If
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", sEnterAsTab)
   
   'Look for EsiExit.exe
   'The user setting may not be correct
   If Dir(sAppPath & "EsiExit.exe") = "" Then
      MsgBox "Cannot Find EsiExit.exe In The Startup Folder." & vbCr _
         & "This File Is A Required Element.  Please Contact " & vbCr _
         & "Your System Administrator.", vbExclamation, sSysCaption
   End
   End If
   If Dir("c:\program Files\ES2000\ES2000.chm") Then
      App.HelpFile = "c:\program Files\ES2000\ES2000.chm"
   Else
      If Dir(sAppPath & "ES2000.hlp") <> "" Then _
             App.HelpFile = sFilePath & "ES2000.hlp"
   End If
   
   'sFilePath = Trim(GetSetting("Esi2000", "System", "FilePath", sFilePath))   'just use app path
   If Trim(sFilePath) = "" Then sFilePath = sAppPath
   SaveSetting "Esi2000", "System", "FilePath", sFilePath
   iSaveCurrent = GetSetting("Esi2000", "mngr", "CurrentBar", iSaveCurrent)
   bShowVertical = GetSetting("Esi2000", "mngr", "ShowVertical", bShowVertical)
   
   'if running in VB, allow report path different from
   If RunningInIDE Then
      sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   End If
   If sReportPath = "" Then
      sReportPath = sAppPath
      'SaveSetting "Esi2000", "System", "ReportPath", sReportPath
   End If
   bUserLoggedOn = 0
   
'   Set RdoEnv = rdoEnvironments(0)
   RegisterUser
   CheckSecurity
   CheckEventLog
   If bRetry = 1 Then CheckSecurity
   EsiLogon.Show
      
End Sub

Public Sub KeyLock(KeyAscii As Integer, frm As Form)
   If KeyAscii <> 13 Then
      If Not bInsertOn Then
         If KeyAscii > 32 Then SendKeys "+{RIGHT}", True
      End If
   End If
   
End Sub


Public Function CloseManager() As Byte

   'return = 0 of closing
   '       = 1 if not closing
   
   Dim iList As Integer
   Dim n As Integer
   Dim msg As String
   Dim lClose As Long
   Dim lStillOpen(14) As Long
   
   On Error Resume Next
   'See if any are active
   CloseManager = 0
   lStillOpen(6) = FindWindow(vbNullString, "ESI Time") '*
   
   lStillOpen(7) = FindWindow(vbNullString, "ESI Administration") '*
   
   lStillOpen(8) = FindWindow(vbNullString, "ESI Sales") '*
   
   lStillOpen(9) = FindWindow(vbNullString, "ESI Engineering") '*
   
   lStillOpen(10) = FindWindow(vbNullString, "ESI Quality") '*
   
   lStillOpen(11) = FindWindow(vbNullString, "ESI Production") '*
   
   lStillOpen(12) = FindWindow(vbNullString, "ESI Inventory") '*
   
   lStillOpen(13) = FindWindow(vbNullString, "ESI Finance") '*
   bByte = 0
   For iList = 0 To 13
      If lStillOpen(iList) > 0 Then bByte = 1
   Next
   
   If bByte = 1 Then
      lClose = FindWindow(vbNullString, "ESI CloseSections")
      If lClose = 0 Then
         msg$ = "You Have Sections Open. You Should Close" & vbCr _
                & "The Open Sections Before Logging Off. " & vbCr & vbCr _
                & "Do You Really Want To Log Off Anyway?"
         iList = MsgBox(msg, ES_NOQUESTION, sSysCaption)
         If iList = vbNo Then
            CloseManager = 1
            Exit Function
         End If
      End If
   End If
   MouseCursor 13
   'Remove from OpenTask
   Alert.Timer1.Enabled = False
   iList = 0
   
   '2/18/05 Close them and remove from stack
   If bByte = 1 Then
      For iList = 0 To 13
         'Old
         If lStillOpen(iList) > 0 Then
            SendMessage lStillOpen(iList), WM_CLOSE, 0&, 0&
            Sleep 500
         End If
      Next
   End If
   iSaveCurrent = GetSetting("Esi2000", "mngr", "CurrentBar", iSaveCurrent)
   SaveSetting "Esi2000", "system", "UserId", ""
   SaveSetting "Esi2000", "Sections", "EsiOpen", 0
   SaveSetting "Esi2000", "System", "UserProfileRec", 0
   SaveSetting "Esi2000", "mngr", "ShowVertical", bShowVertical
   
   CloseUser
   CloseServer
   MouseCursor 0
   On Error Resume Next
   For iList = Forms.Count - 1 To 0
      Unload Forms(iList)
   Next
End

End Function

'reduces 10/22/01

Public Sub SetFormSize(frm As Form, Optional bSkipColors As Byte)
   Dim i As Integer
   On Error Resume Next
   Dim cNewSize As Currency
   cNewSize = 1
   If Screen.Width \ Screen.TwipsPerPixelX = 1024 Then
      cNewSize = 1.1
   ElseIf Screen.Width \ Screen.TwipsPerPixelX = 1280 Then cNewSize = 1.13
   Else
      If Screen.Width \ Screen.TwipsPerPixelX > 1590 Then cNewSize = 1.14
   End If
   frm.Height = frm.Height * cNewSize
   frm.Width = frm.Width * cNewSize
   If bSkipColors = 1 Then Exit Sub
   
   If frm.Name = "EsiLogon" Then Exit Sub
   frm.BackColor = RGB(212, 208, 200)
   For i = 0 To frm.Controls.Count - 1
      If TypeOf frm.Controls(i) Is TextBox Then
         frm.Controls(i).BackColor = vbWhite
         frm.Controls(i).ForeColor = Es_TextForeColor
      ElseIf TypeOf frm.Controls(i) Is CommandButton Then
         frm.Controls(i).BackColor = RGB(212, 208, 200)
         frm.Controls(i).ForeColor = Es_TextForeColor
         
      ElseIf TypeOf frm.Controls(i) Is CheckBox Then
         frm.Controls(i).BackColor = RGB(212, 208, 200)
         If Left$(frm.Controls(i).Caption, 2) = "__" Then
            frm.Controls(i).ForeColor = Es_CheckBoxForeColor
         Else
            frm.Controls(i).ForeColor = Es_TextForeColor
         End If
         
      ElseIf TypeOf frm.Controls(i) Is Frame Or TypeOf frm.Controls(i) Is OptionButton Then
         frm.Controls(i).BackColor = RGB(212, 208, 200)
         frm.Controls(i).ForeColor = Es_TextForeColor
      Else
         If TypeOf frm.Controls(i) Is ComboBox Then
            frm.Controls(i).BackColor = Es_TextBackColor
            frm.Controls(i).ForeColor = Es_TextForeColor
            If frm.Controls(i).Tag <> "" Then
               If frm.Controls(i).Tag = 4 Then
                  frm.Controls(i).ToolTipText = "Date As 09/01/05, " _
                               & "09.01.05, 9/1/05 Or Pull Down"
               End If
            End If
         End If
      End If
   Next
   
End Sub

'Code for updating and adding tables
'Will only work if the user actually logs on as the
'Administrator (should be blind to the user).
'2/6/06 Added registry settings to speed loads


Function Compress(TestNo As Variant) As String
   'Remove spaces (32), dashes (45) and tabs (9) from indexed fields to
   'avoid duplicate entries
   Dim a As Integer
   Dim k As Integer
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo EsiFComp1
   PartNo = Trim$(TestNo)
   a = Len(PartNo)
   If a > 0 Then
      For k = 1 To a
         If Mid$(PartNo, k, 1) <> Chr$(45) And Mid$(PartNo, k, 1) <> Chr$(32) And Mid$(PartNo, k, 1) <> Chr$(9) Then
            NewPart = NewPart & Mid$(PartNo, k, 1)
         End If
      Next
   End If
   Compress = Trim(NewPart)
   Exit Function
   
EsiFComp1:
   Resume EsiFComp2
EsiFComp2:
   On Error Resume Next
   Compress = TestNo
   
End Function


Public Sub CloseServer()
   On Error Resume Next
'   RdoCon.Close
'   RdoEnv.Close
'   Set RdoEnv = Nothing
   
End Sub

Public Function RegisterSqlDsn(sDataSource As String) As String
   'Build keywords string for registering or altering
   'DSN (Datasource Name) for Crystal reports
    Dim sAttribs As String
    If sDataSource = "" Then sDataSource = "ESI2000"
    On Error GoTo modErr1
    sAttribs = "Description=" _
            & "ES/2002 SQL Server Data" _
      & vbCr & "OemToAnsi=No" _
      & vbCr & "SERVER=" & sserver _
      & vbCr & "Database=Esi2000Db"
    ' Create new/revise registered DSN.
'TODO:    rdoEngine.rdoRegisterDataSource sDataSource, _
'             "SQL Server", True, sAttribs
    RegisterSqlDsn = sDataSource
    Exit Function

modErr1:
    RegisterSqlDsn = sDataSource
    On Error GoTo 0

End Function

Public Sub CloseUser()
   Dim sDate As String
   sDate = Format(GetServerDateTime(), "mm/dd/yy hh:mm AM/PM")
   On Error GoTo modErr1
'   sSql = "use msdb"
'   RdoCon.Execute sSql, rdExecDirect
'
'   sSql = "UPDATE SystemUserLog SET Log_Off='" & sDate & "' " _
'          & "WHERE Log_ID=" & lLogId & " "
'   RdoCon.Execute sSql, rdExecDirect
'
'   sSql = "use " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

Public Function GetWorkStation() As String
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Long
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   On Error GoTo modErr1
   
   'Get the workstation and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   '
   If GetComputerName(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetWorkStation = strBuf
   Else
      GetWorkStation = ""
   End If
   Exit Function
   
modErr1:
   On Error GoTo 0
   GetWorkStation = ""
   
End Function

Public Function GetNetId() As String
   Dim intZeroPos As Integer
   Dim gintMAX_SIZE As Long
   Dim strBuf As String
   gintMAX_SIZE = 255 'Maximum buffer size
   
   strBuf = Space$(gintMAX_SIZE)
   '
   'Get the Net user login and then trim the buffer to the exact length
   'returned and add a dir sep (backslash) if the API didn't return one
   
   On Error GoTo modErr1
   If GetUserName(strBuf, gintMAX_SIZE) > 0 Then
      intZeroPos = InStr(strBuf, Chr$(0))
      If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
      GetNetId = strBuf
   Else
      GetNetId = ""
   End If
   Exit Function
   
modErr1:
   On Error GoTo 0
   GetNetId = ""
   
End Function


Public Sub RegisterUser()
   On Error GoTo modErr1
   sWorkStation = GetWorkStation()
   sUserName = GetNetId()
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

Public Sub ModuleNotFound(sModule As String)
   MsgBox "Module " & sModule & " Wasn't Found Or No User   " & vbCr _
      & "Permissions. Check Path, Settings And Permissions.", vbExclamation, sSysCaption
   
End Sub

'Public Function GetSystemCaption() As String
'    GetSystemCaption = "ES/2000 ERP"
'
'End Function
'
'

'3/7/02 for Crystal Reports TestDb

Public Function RegisterTestDbDsn(sDataSrc) As String
   'Build keywords string for registering or altering
   'DSN (Datasource Name) for Crystal reports
   Dim sAttribs As String
   On Error GoTo modErr1
'   sAttribs = "Description=" _
'              & "ES/2002 SQL Server Test Data" _
'              & vbCr & "OemToAnsi=No" _
'              & vbCr & "SERVER=" & sserver _
'              & vbCr & "Database=" & sDataSrc
'   ' Create new/revise registered DSN.
'   rdoEngine.rdoRegisterDataSource sDataSrc, _
'      "SQL Server", True, sAttribs
'   RegisterTestDbDsn = sDataSrc
   Exit Function
   
modErr1:
   RegisterTestDbDsn = sDataSrc
   On Error GoTo 0
   
End Function

'Second Release of Lot Tables
'New 3/25/02
'Add PO Columns and expanded User Lot ID



'3/23/03 Confirm users
'Build a temp table and populate it with unigue values
'compare to Users

Public Function CheckLogons() As Byte
'   Dim RdoLog As rdoResultset
'   Dim b As Byte
'   Dim bUsers As Byte
'   Dim sTable As String
'
'   On Error Resume Next
'   sSql = "use msdb"
'   RdoCon.Execute sSql, rdExecDirect
'
'   'Get the Users. Default is 99
'   sSql = "SELECT ChannelOps from syschannels where ChannelRow=1"
'   bSqlRows = GetDataSet(RdoLog, ES_FORWARD)
'   If bSqlRows Then bUsers = Asc(RdoLog!ChannelOps) - 10
'   If bUsers = 0 Then bUsers = 99
'
'   sSql = "use " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
'
'   sTable = EsiLogon.txtUsr & "Log"
'   If bUsers > 0 And bUsers < 99 Then
'      sSql = "CREATE TABLE " & sTable & " (" _
'             & "HostName CHAR(50) NULL DEFAULT(''))"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "CREATE UNIQUE CLUSTERED INDEX LogRef ON dbo." & sTable & " (HostName) WITH FILLFACTOR = 80"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "sp_who"
'      bSqlRows = GetDataSet(RdoLog)
'      If bSqlRows Then
'         'just in case it didn't get killed
'         sSql = "Truncate TABLE " & sTable & " "
'         RdoCon.Execute sSql, rdExecDirect
'         With RdoLog
'            Do Until .EOF
'               Err = 0
'               If Trim(!HostName) <> "" Then
'                  sSql = "INSERT INTO " & sTable & " (" _
'                         & "HostName) VALUES('" & Trim(!HostName) & "')"
'                  RdoCon.Execute sSql, rdExecDirect
'                  If Err = 0 Then b = b + 1
'               End If
'               .MoveNext
'            Loop
'            .Cancel
'         End With
'      End If
'      sSql = "DROP TABLE " & sTable & " "
'      RdoCon.Execute sSql, rdExecDirect
'   End If
'   If b > bUsers + 1 Then
'      CheckLogons = 1
'   Else
'      CheckLogons = 0
'   End If
   
End Function






Public Sub GetAppTitles()
   '    sAppTitles(0) = GetSetting("Esi2000", "AppTitle", "admn", sAppTitles(0))
   '    If sAppTitles(0) = "" Then sAppTitles(0) = "ESI Administration"
   SetSectionInfo SECTION_ADMN, "Admn", "ESI Administration", "Office Administration", _
      "System Administration" & vbCrLf _
      & "Sales Administration" & vbCrLf _
      & "Production Control Administration" & vbCrLf _
      & "Time Charges" & vbCrLf _
      & "Inventory Management"
   '
   '    sAppTitles(3) = GetSetting("Esi2000", "AppTitle", "prod", sAppTitles(3))
   '    If sAppTitles(3) = "" Then sAppTitles(3) = "ESI Production"
   SetSectionInfo SECTION_PROD, "Prod", "ESI Production", "Production", _
      "Inventory Control" & vbCr _
      & "Purchasing/Receiving " & vbCr _
      & "Material Control" & vbCr _
      & "Capacity Requirements Planning" & vbCr _
      & "Shop Floor Control" & vbCr _
      & "Receiving"
   
   '
   '    sAppTitles(2) = GetSetting("Esi2000", "AppTitle", "engr", sAppTitles(2))
   '    If sAppTitles(2) = "" Then sAppTitles(2) = "ESI Engineering"
   SetSectionInfo SECTION_ENGR, "Engr", "ESI Engineering", "Engineering", _
      "Estimating" & vbCr _
      & "Bills of Material" & vbCrLf _
      & "Product Routings" & vbCrLf _
      & "Product Costing" & vbCrLf _
      & "Document Control" & vbCrLf _
      & "Tooling Management"
   
   '
   '    sAppTitles(1) = GetSetting("Esi2000", "AppTitle", "sale", sAppTitles(1))
   '    If sAppTitles(1) = "" Then sAppTitles(1) = "ESI Sales"
   SetSectionInfo SECTION_SALE, "Sale", "ESI Sales", "Sales", _
      "Sales Processing" & vbCrLf _
      & "Packing Slip Management " & vbCrLf _
      & "Bookings and Backlog"
   
   '
   '    sAppTitles(6) = GetSetting("Esi2000", "AppTitle", "fina", sAppTitles(6))
   '    If sAppTitles(6) = "" Then sAppTitles(6) = "ESI Finance"
   SetSectionInfo SECTION_FINA, "Fina", "ESI Finance", "Financial Accounting", _
      "Accounts Receivable" & vbCrLf _
      & "Customer Credit Management" & vbCrLf _
      & "Accounts Payable"
   
   '    sAppTitles(5) = GetSetting("Esi2000", "AppTitle", "qual", sAppTitles(5))
   '    If sAppTitles(5) = "" Then sAppTitles(5) = "ESI Quality"
   SetSectionInfo SECTION_QUAL, "Qual", "ESI Quality", "Quality Assurance", _
      "Rejection Tag Management" & vbCrLf _
      & "First Article Inspection" & vbCrLf _
      & "Statistical Process Control" & vbCrLf _
      & "Component Traceability" & vbCrLf _
      & "On Dock Inspection"
   
   '
   '    sAppTitles(4) = GetSetting("Esi2000", "AppTitle", "invc", sAppTitles(4))
   '    If sAppTitles(4) = "" Then sAppTitles(4) = "ESI Inventory"
   SetSectionInfo SECTION_INVC, "Invc", "ESI Inventory", "Inventory Control", _
      "Inventory Control" & vbCrLf _
      & "Purchasing Information" & vbCrLf _
      & "Material Control" & vbCrLf _
      & "Inventory Management"
   
   SetSectionInfo SECTION_POM, "POM", "Point Of Manufacturing", "Point Of Manufacturing", ""
   
   SetSectionInfo SECTION_TIME, "Time", "ESI Time Management", "Time Management", _
      "Time Entry"
   
   SetSectionInfo SECTION_DATACOL, "DataCollection", "Data Collection", "DataCollection", _
      "DataCollection"
      
End Sub

Private Sub SetSectionInfo(SectionNo As Integer, Title As String, DefaultName As String, _
                           Description As String, Buttons As String)
   SectionInfo(SectionNo).AppTitle = Title
   SectionInfo(SectionNo).AppName = GetSetting("Esi2000", "AppTitle", Title, DefaultName)
   SectionInfo(SectionNo).ExeName = "Esi" & Title & ".exe"
   SectionInfo(SectionNo).AppDescription = Description
   SectionInfo(SectionNo).AppButtons = Buttons
End Sub

'3/21/05 Make the Error.Log header if it doesn't exist
'11/1/05 Eliminated Event log (replaced by msdb log)

Private Sub CheckEventLog()
   '    Dim iFreeFile As Integer
   '    Dim iFileLen  As Integer
   '
   '    Dim sDate       As String * 16
   '    Dim sSection    As String * 8
   '    Dim sForm       As String * 12
   '    Dim sErrNum     As String * 10
   '    Dim sErrSev     As String * 2
   '    Dim sProc       As String * 10
   '    Dim sUserName   As String * 20
   '        If Dir(sFilePath & "EsiError.log") = "" Then
   '            iFreeFile = FreeFile
   '            Open sFilePath & "EsiError.log" For Append Shared As iFreeFile
   '            iFileLen = LOF(iFreeFile)
   '                If iFileLen = 0 Then
   '                    sDate = "Date"
   '                    sSection = "Section"
   '                    sForm = "Form"
   '                    sUserName = "User"
   '                    sErrNum = "Err"
   '                    sProc = "Procedure"
   '                    Print #iFreeFile, "System Event Log"
   '                    Print #iFreeFile, "Note: Information messages and User input notices are not recorded"
   '                    Print #iFreeFile, sDate; "    "; sSection; "  "; sForm; "  "; sUserName; _
   '                      "          "; sErrNum; "   "; sErrSev; " "; sProc; " "
   '                    Print #iFreeFile, String(156, "-")
   '                End If
   '            Close iFreeFile
   '        End If
   
End Sub

'7/13/05 Delete obsolete help

Public Sub KillOldHelp()
   '    Dim sHelpPath As String
   '    sHelpPath = App.Path & "\"
   '    If Dir(sHelpPath & "ESIFINA.HLP") <> "" Then _
   '            Kill sHelpPath & "ESIFINA.HLP"
   '    If Dir(sHelpPath & "ESIFINA.CNT") <> "" Then _
   '            Kill sHelpPath & "ESIFINA.CNT"
   '
   '    If Dir(sHelpPath & "ESIPROD.HLP") <> "" Then _
   '            Kill sHelpPath & "ESIPROD.HLP"
   '    If Dir(sHelpPath & "ESIPROD.CNT") <> "" Then _
   '            Kill sHelpPath & "ESIPROD.CNT"
   '
   '    If Dir(sHelpPath & "ESIADMN.HLP") <> "" Then _
   '            Kill sHelpPath & "ESIADMN.HLP"
   '    If Dir(sHelpPath & "ESIADMN.CNT") <> "" Then _
   '            Kill sHelpPath & "ESIADMN.CNT"
   '
   '    If Dir(sHelpPath & "ESIENGR.HLP") <> "" Then _
   '            Kill sHelpPath & "ESIENGR.HLP"
   '    If Dir(sHelpPath & "ESIENGR.CNT") <> "" Then _
   '            Kill sHelpPath & "ESIENGR.CNT"
   '
   '    If Dir(sHelpPath & "ESIQUAL.HLP") <> "" Then _
   '            Kill sHelpPath & "ESIQUAL.HLP"
   '    If Dir(sHelpPath & "ESIQUAL.CNT") <> "" Then _
   '            Kill sHelpPath & "ESIQUAL.CNT"
   '
   '    If Dir(sHelpPath & "ESISALE.HLP") <> "" Then _
   '            Kill sHelpPath & "ESISALE.HLP"
   '    If Dir(sHelpPath & "ESISALE.CNT") <> "" Then _
   '            Kill sHelpPath & "ESISALE.CNT"
   
End Sub

'7/14/05 Install Local Help
'7/28/05 Revised

Public Sub CopyHelp()
   'Dim b          As Byte
   Dim sHelpPath As String
   Dim sFileDate As String
   Dim sPriorDate As String
   Dim sWindows As String
   
   On Error Resume Next
   '    b = RunningInIDE()
   '        If b = 0 Then
   '            If sHelpPath = "" Then sHelpPath = App.Path & "\"
   '        Else
   '            sHelpPath = "c:\esi2000\"
   '        End If
   
   sHelpPath = App.Path & "\" '8/3/07 always run in app dir
   
   sFileDate = Format(FileDateTime(sHelpPath & "ES2000.chm"), "yyyy-mm-dd")
   Err.Clear
   sPriorDate = Format(FileDateTime("c:\Program Files\ES2000\ES2000.chm"), "yyyy-mm-dd")
   If Err > 0 Then sPriorDate = "2005-01-01"
   Err.Clear
   'Has it been loaded? Is there a newer one?
   If Dir("c:\Program Files\ES2000") = "" Then MkDir "c:\Program Files\ES2000"
   Err.Clear
   If sFileDate > sPriorDate Then
      If Dir(sHelpPath & "ES2000.chm") <> "" Then
         FileCopy sHelpPath & "ES2000.chm", "c:\Program Files\ES2000\ES2000.chm"
      End If
   End If
   
   '8/14/06/8/17/06 According to Steve W, we won't need this
   '    sWindows = GetWindowsDir() & "\system32\"
   '     If Dir(sWindows & "MSCOMCTL.OCX") = "" Then _
   '        FileCopy App.Path & "\" & "MSCOMCTL.OCX", sWindows & "MSCOMCTL.OCX"
   
End Sub

'7/27/05 blocked 2/6/06

Public Sub CreateSplitLots()
   '    Dim RdoSplit As rdoResultset
   '    On Error Resume Next
   '    sSql = "SELECT SPLIT_FROMLOT FROM LospTable WHERE SPLIT_FROMLOT='FUBAR'"
   '    RdoCon.Execute sSql, rdExecDirect
   '        If Err > 0 Then
   '            Err.Clear
   '            sSql = "CREATE TABLE LospTable (" _
   '                & "SPLIT_FROMLOT CHAR(15) NULL DEFAULT('')," _
   '                & "SPLIT_TOLOT CHAR(15) NULL DEFAULT('')," _
   '                & "SPLIT_DATE SMALLDATETIME NULL DEFAULT(GetDate())," _
   '                & "SPLIT_QUANTITY REAL NULL DEFAULT(0))"
   '            RdoCon.Execute sSql, rdExecDirect
   '            If Err = 0 Then
   '                Err.Clear
   '                sSql = "CREATE UNIQUE CLUSTERED INDEX SplitRef ON " _
   '                    & "LospTable(SPLIT_FROMLOT,SPLIT_TOLOT) WITH FILLFACTOR = 80"
   '                RdoCon.Execute sSql, rdExecDirect
   '            End If
   '            If Err = 0 Then
   '                sSql = "SELECT LOTNUMBER,LOTADATE,LOTORIGINALQTY,LOTSPLITFROMSYS " _
   '                    & "FROM LohdTable WHERE LOTSPLITFROMSYS<>''"
   '                bSqlRows = GetDataSet(RdoSplit, ES_FORWARD)
   '                If bSqlRows Then
   '                    With RdoSplit
   '                        Do Until .EOF
   '                            sSql = "INSERT INTO LospTable (SPLIT_FROMLOT," _
   '                                & "SPLIT_TOLOT,SPLIT_DATE,SPLIT_QUANTITY) " _
   '                                & "VALUES('" & !LOTSPLITFROMSYS & "','" _
   '                                & Trim(!LOTNUMBER) & "','" & !LOTADATE & "'," _
   '                                & !LOTORIGINALQTY & ")"
   '                            RdoCon.Execute sSql, rdExecDirect
   '                            .MoveNext
   '                        Loop
   '                        .Cancel
   '                        Do Until .MoreResults = 0
   '                        Loop
   '                    End With
   '                End If
   '            End If
   '        End If
   
End Sub

Public Sub GetSystemMessage()
'   Dim RdoMsg As rdoResultset
'   Dim lMsgID As Long
'   Dim lOldId As Long
'   On Error Resume Next
'   RdoCon.Execute "use msdb"
'   sSql = "SELECT MAX(Message_ID) as ID FROM SystemMessages"
'   bSqlRows = GetDataSet(RdoMsg, ES_FORWARD)
'   If bSqlRows Then
'      If Not IsNull(RdoMsg!ID) Then
'         lMsgID = RdoMsg!ID
'      Else
'         lMsgID = 0
'      End If
'   Else
'      lMsgID = 0
'   End If
'   If lMsgID > 0 Then
'      Err.Clear
'      lOldId = GetSetting("Esi2000", "messages", "lastmessage", lOldId)
'      If lMsgID > lOldId Then
'         Err.Clear
'         sSql = "SELECT * FROM SystemMessages WHERE Message_ID=" & lMsgID
'         bSqlRows = GetDataSet(RdoMsg, ES_FORWARD)
'         If bSqlRows Then
'            With RdoMsg
'               SysMessage.z1(3).Caption = "You Have A New Message."
'               SysMessage.txtTime = "" & Trim(!Message_TIME)
'               SysMessage.txtFrom = "" & Trim(!Message_FROM)
'               SysMessage.txtSubject = "" & Trim(!Message_HEADER)
'               SysMessage.txtMessage = "" & Trim(!Message_TEXT)
'               SysMessage.Show
'               .Cancel
'            End With
'         End If
'         If Err > 0 Then
'            MsgBox "You Have A Message That Can't Be Read" & vbCr _
'               & "Permissions To msdb Is Required.  Please" & vbCr _
'               & "Contact Your Systems Administrator.", _
'               vbInformation, sSysCaption
'            RdoCon.Execute "use " & sDataBase
'            Exit Sub
'         End If
'      End If
'   End If
'   SaveSetting "Esi2000", "messages", "lastmessage", lMsgID
'   Set RdoMsg = Nothing
'   RdoCon.Execute "use " & sDataBase
   
End Sub


Private Sub CreateSystemUserLog()
'   Dim sAttribs As String
'   On Error Resume Next
'   sSql = "use msdb"
'   RdoCon.Execute sSql, rdExecDirect
'
'   sSql = "select Log_ID from SystemUserLog WHERE Log_ID=1"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err > 0 Then
'      Err.Clear
'      sSql = "CREATE TABLE SystemUserLog (" _
'             & "Log_ID int IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL," _
'             & "Log_On char(18) NULL DEFAULT('')," _
'             & "Log_Off char(18) NULL DEFAULT('')," _
'             & "Log_Workstation char(20) NULL DEFAULT('')," _
'             & "Log_User char(20) NULL DEFAULT('')," _
'             & "Log_Path char(40) NULL DEFAULT('')," _
'             & "Log_DataBase char(30) NULL DEFAULT(''))"
'      RdoCon.Execute sSql, rdExecDirect
'      If Err = 0 Then
'         sSql = "CREATE UNIQUE CLUSTERED INDEX EventRow ON " _
'                & "SystemUserLog(Log_ID) WITH FILLFACTOR = 80"
'         RdoCon.Execute sSql, rdExecDirect
'      End If
'   End If
'   sSql = "use " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
   
End Sub

Public Sub WriteUserLog()
'   Dim RdoLog As rdoResultset
'   Dim sDate As String
'   On Error Resume Next
'   sSql = "use msdb"
'   RdoCon.Execute sSql, rdExecDirect
'
'   'Clean Out Trash
'   sSql = "DELETE FROM SystemUserLog WHERE (Log_User='" & Trim(EsiLogon.txtUsr) _
'          & "' AND Log_Off='Active' AND Log_Workstation='" & sWorkStation & "')"
'   RdoCon.Execute sSql, rdExecDirect
'
'   sDate = Format(GetServerDateTime(), "mm/dd/yy hh:mm AM/PM")
'   sSql = "SELECT * FROM SystemUserLog WHERE Log_ID=1"
'   bSqlRows = GetDataSet(RdoLog, ES_KEYSET)
'   With RdoLog
'      .AddNew
'      !Log_On = sDate
'      !Log_Off = "Active"
'      !Log_Workstation = sWorkStation
'      !Log_User = Trim(EsiLogon.txtUsr)
'      !Log_Path = sFilePath
'      !Log_Database = sDataBase
'      .Update
'      .Cancel
'   End With
'   sSql = "SELECT MAX(Log_ID) as LogId FROM SystemUserLog "
'   bSqlRows = GetDataSet(RdoLog, ES_FORWARD)
'   lLogId = RdoLog!LogId
'   If lLogId > 100 Then
'      sSql = "DELETE FROM SystemUserLog WHERE (Log_Off='Active' AND " _
'             & "Log_ID<" & (lLogId - 50) & ") "
'      RdoCon.Execute sSql, rdExecDirect
'   End If
'   RdoLog.Close
'   Set RdoLog = Nothing
'   sSql = "use " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
   
End Sub



Public Sub DropWipAccounts()
'   Dim RdoDrop As rdoResultset
'   Dim Cl As rdoColumn
'   Dim sConstraint As String
'
'   On Error Resume Next
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPCRLABACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPCRLABACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPCRLABACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPCRMATACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPCRMATACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPCRMATACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPCREXPACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPCREXPACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPCREXPACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPCROHDACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPCROHDACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPCROHDACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   '****
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPDRLABACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPDRLABACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPDRLABACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPDRMATACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPDRMATACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPDRMATACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPDREXPACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPDREXPACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPDREXPACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'
'   sSql = "sp_columns @table_name=ComnTable,@column_name=WIPDROHDACCT1"
'   Set RdoDrop = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
'   With RdoDrop
'      sSql = "ALTER TABLE ComnTable DROP DF_ComnTable_WIPDROHDACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE ComnTable DROP COLUMN WIPDROHDACCT1"
'      RdoCon.Execute sSql, rdExecDirect
'   End With
'   Err.Clear
'   Set RdoDrop = Nothing
'
End Sub

'5/3/06 Telecon/Email to Larry verifying inclusion in 5/8/06 Update


Public Function CheckConvErrors() As Byte
   CheckConvErrors = 0
   'TODO: ADO Errors
'   For Each ER In rdoErrors
'      If Left(ER.Description, 5) = "22003" Then
'         iColCounter = iColCounter + 1
'         sBadColumns(iColCounter) = RdoCol.rdoColumns(2) & "." & RdoCol.rdoColumns(3)
'         CheckConvErrors = 1
'      End If
'   Next ER
   
End Function


'5/15/06 Primary/Foreign Keys
'Not all tables are included. This tactic precludes entering PO Items, SO Items
'and Bids (for example) without a valid Part Number (we handle that).
'Inventory Activity would halt deleting any Part Number entered in error.

Private Sub KeysPartTable()
'   On Error Resume Next
'   Err.Clear
'   sSql = "SELECT PartKeys FROM Preferences WHERE PreRecord=1"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err > 0 Then
'      Err.Clear
'      'Clear bogus
'      sSql = "DELETE FROM InvaTable " & vbCr _
'             & "FROM InvaTable LEFT JOIN PartTable ON InvaTable.INPART = PartTable.PARTREF " & vbCr _
'             & "WHERE (PartTable.PARTREF Is Null)"
'      RdoCon.Execute sSql, rdExecDirect
'
'      Err.Clear
'      sSql = "ALTER TABLE Preferences ADD PartKeys smalldatetime null"
'      RdoCon.Execute sSql, rdExecDirect
'
'      Err.Clear
'      RdoCon.BeginTrans
'      sSql = "DROP INDEX PartTable.PartRef"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE PartTable ADD Constraint PK_PartTable_PARTREF PRIMARY KEY CLUSTERED (PARTREF) " _
'             & "WITH FILLFACTOR=80 "
'      RdoCon.Execute sSql, rdExecDirect
'      If Err = 0 Then
'         RdoCon.CommitTrans
'      Else
'         RdoCon.RollbackTrans
'         GoTo PKeys
'      End If
'
'      sSql = "ALTER TABLE RunsTable ADD CONSTRAINT FK_RunsTable_PartTable FOREIGN KEY (RUNREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE RnopTable ADD CONSTRAINT FK_RnopTable_PartTable FOREIGN KEY (OPREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE "
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE BmhdTable ADD CONSTRAINT FK_BmhdTable_PartTable FOREIGN KEY (BMHREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE "
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE BmplTable ADD CONSTRAINT FK_BmplTable_PartTable FOREIGN KEY (BMASSYPART) References PartTable ON DELETE CASCADE ON UPDATE CASCADE"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE PsitTable ADD CONSTRAINT FK_PsitTable_PartTable FOREIGN KEY (PIPART) References PartTable ON DELETE CASCADE ON UPDATE CASCADE"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE LohdTable ADD CONSTRAINT FK_LohdTable_PartTable FOREIGN KEY (LOTPARTREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE "
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE LoitTable ADD CONSTRAINT FK_LoitTable_PartTable FOREIGN KEY (LOIPARTREF) References PartTable ON DELETE CASCADE ON UPDATE CASCADE "
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "ALTER TABLE InvaTable ADD CONSTRAINT FK_InvaTable_PartTable FOREIGN KEY (INPART) References PartTable ON DELETE CASCADE ON UPDATE CASCADE"
'      RdoCon.Execute sSql, rdExecDirect
'
'      'Debilitating
'      'sSql = "ALTER TABLE PoitTable ADD CONSTRAINT FK_PoitTable_PartTable FOREIGN KEY (PIPART) References PartTable"
'      'RdoCon.Execute sSql, rdExecDirect
'
'      'sSql = "ALTER TABLE SoitTable ADD CONSTRAINT FK_SoitTable_PartTable FOREIGN KEY (ITPART) References PartTable"
'      'RdoCon.Execute sSql, rdExecDirect
'      'Most Finance Tables might have the same problem
'
'PKeys:
'      sSql = "UPDATE Preferences SET PartKeys='" & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE PreRecord=1"
'      RdoCon.Execute sSql, rdExecDirect
'   End If
   
End Sub

'5/16/06 New or Rebuild Triggers in the event that they were deleted

Private Sub BuildTriggers()
   Dim bCounter As Byte
   On Error Resume Next
   
'   bCounter = Val(GetSetting("Esi2000", "System", "Triggers1", bCounter))
'   If bCounter = 0 Then
'      sSql = "use msdb"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/16/06 Delete Trigger (syschannels)
'      sSql = "CREATE TRIGGER DT_syschannels ON syschannels" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      sSql = "use " & sDataBase
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/16/06 Delete Trigger (Runs)
'      sSql = "CREATE TRIGGER DT_RunsTable ON RunsTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/16/06 Delete Trigger (PO's)
'      sSql = "CREATE TRIGGER DT_PohdTable ON PohdTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SavePoRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/16/06 Delete Trigger (SO's)
'      sSql = "CREATE TRIGGER DT_SohdTable ON SohdTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveSoRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/16/06 Delete Trigger (PS's)
'      sSql = "CREATE TRIGGER DT_PshdTable ON PshdTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SavePsRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/17/06 Delete Trigger (Comn)
'      sSql = "CREATE TRIGGER DT_ComnTable ON ComnTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'      Err.Clear
'      '5/17/06 Delete Trigger (Pref)
'      sSql = "CREATE TRIGGER DT_Preferences ON Preferences" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'
'      '5/17/06 Delete Trigger (Empl)
'      sSql = "CREATE TRIGGER DT_EmplTable ON EmplTable" & vbCr _
'             & "FOR  DELETE " & vbCr _
'             & "  AS " & vbCr _
'             & "  SAVE TRANSACTION SaveRows " & vbCr _
'             & "  Rollback TRANSACTION"
'      RdoCon.Execute sSql, rdExecDirect
'      Err.Clear
'   End If
   
   bCounter = bCounter + 1
   If bCounter = 15 Then bCounter = 0
   SaveSetting "Esi2000", "System", "Triggers1", bCounter
   
End Sub

'5/24/06 Parts, Inva, Lots
'See email sent to Larry Heck at page bottom 5/24/06

'5/31/06

'10/6/06 Moved Here

'Private Sub InstallSysChannels()
'   On Error Resume Next
'   sSql = "use msdb"
'   RdoCon.Execute sSql, rdExecDirect
'   Err.Clear
'   sSql = "SELECT ChannelOps FROM syschannels WHERE ChannelOps='f'"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err > 0 Then
'      Err.Clear
'      'Sec1 = Admn
'      'Sec2 = Sales
'      'Sec3 = Engr
'      'Sec4 = Prod
'      'Sec5 = Invc
'      'Sec6 = Qual
'      'Sec7 = Fina
'      'Sec8 = Spare
'      sSql = "CREATE TABLE syschannels (" _
'             & "ChannelCreate SMALLDATETIME NULL," _
'             & "ChannelRevise SMALLDATETIME NULL," _
'             & "ChannelOps CHAR(1) NULL DEFAULT('')," _
'             & "Channel1a CHAR(1) NULL DEFAULT('o')," _
'             & "Channel1b CHAR(1) NULL DEFAULT('r')," _
'             & "Channel1c CHAR(1) NULL DEFAULT('v')," _
'             & "Channel1d CHAR(1) NULL DEFAULT('a')," _
'             & "Channel1e CHAR(1) NULL DEFAULT('u')," _
'             & "Channel1f CHAR(1) NULL DEFAULT('s')," _
'             & "Channel1g CHAR(1) NULL DEFAULT('c')," _
'             & "Channel1h CHAR(1) NULL DEFAULT('d'),"
'      sSql = sSql _
'             & "Channel2a CHAR(1) NULL DEFAULT('l')," _
'             & "Channel2b CHAR(1) NULL DEFAULT('g')," _
'             & "Channel2c CHAR(1) NULL DEFAULT('k')," _
'             & "Channel2d CHAR(1) NULL DEFAULT('e')," _
'             & "Channel2e CHAR(1) NULL DEFAULT('j')," _
'             & "Channel2f CHAR(1) NULL DEFAULT('f')," _
'             & "Channel2g CHAR(1) NULL DEFAULT('b')," _
'             & "Channel2h CHAR(1) NULL DEFAULT('i'),"
'      sSql = sSql _
'             & "Channel3a CHAR(1) NULL DEFAULT('')," _
'             & "Channel3b CHAR(1) NULL DEFAULT('')," _
'             & "Channel3c CHAR(1) NULL DEFAULT('')," _
'             & "Channel3d CHAR(1) NULL DEFAULT('')," _
'             & "Channel3e CHAR(1) NULL DEFAULT('')," _
'             & "Channel3f CHAR(1) NULL DEFAULT('')," _
'             & "Channel3g CHAR(1) NULL DEFAULT('')," _
'             & "Channel3h CHAR(1) NULL DEFAULT(''),"
'      sSql = sSql _
'             & "Channel4a CHAR(1) NULL DEFAULT('')," _
'             & "Channel4b CHAR(1) NULL DEFAULT('')," _
'             & "Channel4c CHAR(1) NULL DEFAULT('')," _
'             & "Channel4d CHAR(1) NULL DEFAULT('')," _
'             & "Channel4e CHAR(1) NULL DEFAULT('')," _
'             & "Channel4f CHAR(1) NULL DEFAULT('')," _
'             & "Channel4g CHAR(1) NULL DEFAULT('')," _
'             & "Channel4h CHAR(1) NULL DEFAULT(''),"
'      sSql = sSql _
'             & "Channel5a CHAR(1) NULL DEFAULT('z')," _
'             & "Channel5b CHAR(1) NULL DEFAULT('')," _
'             & "Channel5c CHAR(1) NULL DEFAULT('')," _
'             & "Channel5d CHAR(1) NULL DEFAULT('y')," _
'             & "Channel5e CHAR(1) NULL DEFAULT('')," _
'             & "Channel5f CHAR(1) NULL DEFAULT('')," _
'             & "Channel5g CHAR(1) NULL DEFAULT('')," _
'             & "Channel5h CHAR(1) NULL DEFAULT(''),"
'      sSql = sSql _
'             & "Channel6a CHAR(1) NULL DEFAULT('')," _
'             & "Channel6b CHAR(1) NULL DEFAULT('')," _
'             & "Channel6c CHAR(1) NULL DEFAULT('')," _
'             & "Channel6d CHAR(1) NULL DEFAULT('')," _
'             & "Channel6e CHAR(1) NULL DEFAULT('')," _
'             & "Channel6f CHAR(1) NULL DEFAULT('')," _
'             & "Channel6g CHAR(1) NULL DEFAULT('')," _
'             & "Channel6h CHAR(1) NULL DEFAULT(''),"
'      sSql = sSql _
'             & "Channel7a CHAR(1) NULL DEFAULT('')," _
'             & "Channel7b CHAR(1) NULL DEFAULT('')," _
'             & "Channel7c CHAR(1) NULL DEFAULT('')," _
'             & "Channel7d CHAR(1) NULL DEFAULT('')," _
'             & "Channel7e CHAR(1) NULL DEFAULT('')," _
'             & "Channel7f CHAR(1) NULL DEFAULT('')," _
'             & "Channel7g CHAR(1) NULL DEFAULT('')," _
'             & "Channel7h CHAR(1) NULL DEFAULT(''),"
'      sSql = sSql _
'             & "Channel8a CHAR(1) NULL DEFAULT('')," _
'             & "Channel8b CHAR(1) NULL DEFAULT('')," _
'             & "Channel8c CHAR(1) NULL DEFAULT('')," _
'             & "Channel8d CHAR(1) NULL DEFAULT('')," _
'             & "Channel8e CHAR(1) NULL DEFAULT('')," _
'             & "Channel8f CHAR(1) NULL DEFAULT('')," _
'             & "Channel8g CHAR(1) NULL DEFAULT('')," _
'             & "Channel8h CHAR(1) NULL DEFAULT('')," _
'             & "ChannelId CHAR(10)NULL DEFAULT('')," _
'             & "ChannelRow TINYINT NULL DEFAULT(1))"
'      RdoCon.Execute sSql, rdExecDirect
'
'      If Err = 0 Then
'         sSql = "CREATE UNIQUE CLUSTERED INDEX CnlRow ON " _
'                & "syschannels(ChannelRow) WITH FILLFACTOR = 80"
'         RdoCon.Execute sSql, rdExecDirect
'
'         sSql = "INSERT INTO syschannels (ChannelOps) " _
'                & "VALUES('m')"
'         RdoCon.Execute sSql, rdExecDirect
'      End If
'   End If
'   sSql = "use " & sDataBase
'   RdoCon.Execute sSql, rdExecDirect
'   SaveSetting "ES2000", "system", "channels", 1
'   Err.Clear
'
'End Sub
'
'
'Private Sub InstallSysMessages()
'   On Error Resume Next
'   RdoCon.Execute "use msdb"
'   sSql = "SELECT Message_ID FROM SystemMessages WHERE Message_ID=0"
'   RdoCon.Execute sSql, rdExecDirect
'   If Err > 0 Then
'      Err.Clear
'      sSql = "CREATE TABLE SystemMessages (" _
'             & "Message_ID INT NULL DEFAULT(0)," _
'             & "Message_TIME CHAR(20) NULL DEFAULT('')," _
'             & "Message_FROM CHAR(20) NULL DEFAULT('')," _
'             & "Message_HEADER CHAR(30) NULL DEFAULT('')," _
'             & "Message_TEXT VARCHAR(255) NULL DEFAULT(''))"
'      RdoCon.Execute sSql, rdExecDirect
'      If Err = 0 Then
'         sSql = "CREATE UNIQUE CLUSTERED INDEX MessageRef ON " _
'                & "SystemMessages(Message_ID) WITH FILLFACTOR = 80"
'         RdoCon.Execute sSql, rdExecDirect
'      End If
'   End If
'   RdoCon.Execute "use " & sDataBase
'   SaveSetting "ES2000", "system", "systemmsgs", 1
'
'End Sub
Public Sub UpdateTables()
   MouseCursor 13
'   Dim ER As rdoError
'   Dim Cl As rdoColumn
   
   Dim a As Integer
   Dim C As Integer
   
   Dim sConstraint As String
   
   PopMsg.msg.Alignment = 0
   PopMsg.msg = "Updating Parameters"
   PopMsg.msg.Refresh
   
   'Movement to OldUpdates 1/22/05,2/6/06, 10/6/06 (not all)
   
   On Error Resume Next
   'New Setup establishes first contact
   'sSql = "INSERT INTO Preferences (PreRecord) VALUES(1)"
   sSql = "if not exists (select * from preferences where PreRecord = 1)" & vbCrLf _
    & "INSERT INTO Preferences (PreRecord) VALUES(1)"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   sSql = "UPDATE Preferences SET TimeFormat='##0.000' WHERE (PreRecord=1 AND " _
          & "TimeFormat IS NULL)"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   '8/15/05 to Admin 2/6/02
   '   a = GetSetting("ES2000", "system", "eventlog", a)
   '   If a = 0 Then CreateSystemEventLog
   '   SaveSetting "ES2000", "system", "eventlog", 1
   
'   If Not IsNull(Secure.UserLevel) Then
'      If Secure.UserLevel < 11 Then BuildTriggers
'   Else
'      BuildTriggers
'   End If
   
'   a = GetSetting("ES2000", "system", "channels", a)
'   If a = 0 Then InstallSysChannels
'   PopMsg.msg = "Updating Parameters.."
'   PopMsg.msg.Refresh
   
   '    Err.Clear
   '   2/6/07 To Admin
   '   'Shipped the empty table this code should stay in
   '    CustomReports
   '    a = GetSetting("ES2000", "system", "customreports", a)
   '    If a = 0 Then InstallCustomReports
   Err.Clear
   
   '2/6/07, 5/17/05 Reviewed. 10/12/04 Leave in
'   sSql = "UPDATE ComnTable SET CURPSNUMBER='000000' WHERE COREF=1 AND CURPSNUMBER='' " _
'          & "OR CURPSNUMBER IS NULL"
'   RdoCon.Execute sSql, rdExecDirect
   
   PopMsg.msg = "Updating Parameters..."
   PopMsg.msg.Refresh
   '   6/9/06
   '   '6/23/05 Leave in 10/30/04 Create Default
   '    sSql = "CREATE DEFAULT DEFEMPTY AS ''"
   '    RdoCon.Execute sSql, rdExecDirect
   '
   '   '6/23/05 Leave in 6/28/04 Create Default
   '    sSql = "CREATE DEFAULT DEFZERO AS 0"
   '    RdoCon.Execute sSql, rdExecDirect
   
   PopMsg.msg = "Updating Parameters..."
   PopMsg.msg.Refresh
   
   '2/6/06 Blocked KillOldHelp
   '2/6/06 Blocked CreateSplitLots 'Table
   
   Err.Clear
   '8/10/05 System Messages
'   a = GetSetting("ES2000", "system", "systemmsgs", a)
'   If a = 0 Then InstallSysMessages
   
   '8/26/05 User logging leave in
   CreateSystemUserLog
   SaveSetting "ES2000", "system", "systemmuserlog", 1
   
   Err.Clear
   PopMsg.msg = "Updating Parameters....."
   PopMsg.msg.Refresh
   
   'Continued...
'   sSql = "SELECT InvtoDecimalConvDate FROM Preferences WHERE PreRecord=1 "
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_FORWARD)
'   If bSqlRows Then
'      With RdoTest
'         If IsNull(!InvtoDecimalConvDate) Then ConvertInvColumns
'         .Cancel
'      End With
'   End If
   '*****
   '   2/6/02
   '   'Primary/Foreign Keys (Parts)
   '   '5/30 - 6/9/06
   '    Err.Clear
   '    KeysPartTable
   '    KeysCustomerTables
   '    KeysVendorTables
   '    KeyFinaTables
   
   '10/19/06
   'Electronic Invoicing
   Err.Clear
   sSql = "SELECT CUEINVOICING FROM CustTable WHERE CUREF='fubar'"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   If Err > 0 Then
      Err.Clear
      sSql = "ALTER TABLE CustTable ADD CUEINVOICING TINYINT NULL DEFAULT(0)"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      If Err = 0 Then
         sSql = "UPDATE CustTable SET CUEINVOICING=0 WHERE CUEINVOICING IS NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
   End If
   '12/6/06 Garrett and others
   '    Err.Clear
   '    sSql = "alter table ViitTable drop constraint FK_ViitTable_VihdTable"
   '    RdoCon.Execute sSql, rdExecDirect
   
   '1/8/07 Email to Vendor AR and Customer AP 7.1.3
   Err.Clear
   sSql = "SELECT CUAPEMAIL FROM CustTable WHERE CUAPEMAIL='fubar'"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   If Err > 0 Then
      Err.Clear
      sSql = "ALTER TABLE CustTable ADD CUAPEMAIL VARCHAR(60) NULL DEFAULT('')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      sSql = "ALTER TABLE VndrTable ADD VEAREMAIL VARCHAR(60) NULL DEFAULT('')"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      If Err = 0 Then
         sSql = "UPDATE CustTable SET CUAPEMAIL='' WHERE CUAPEMAIL IS NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
         
         sSql = "UPDATE VndrTable SET VEAREMAIL='' WHERE VEAREMAIL IS NULL"
         clsADOCon.ExecuteSql sSql 'rdExecDirect
      End If
      
   End If
   
   '3/22/07 House Cleaning Safe to remove after update
   sSql = "sp_bindefault DEFZERO, 'LoitTable.LOIMORUNNO'"
   sSql = "UPDATE LoitTable SET LOIMORUNNO=0 WHERE LOIMORUNNO IS NULL"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   PopMsg.msg = "Updating Parameters......"
   PopMsg.msg.Refresh
   
   GoTo ManagerErr2
   Exit Sub
   
ManagerErr1:
   Resume ManagerErr2
ManagerErr2:
   On Error Resume Next
   Set RdoTest = Nothing
   Set RdoCol = Nothing
   'We were just testing and establishing ourselves.
   PopMsg.msg = "End Parameter Updates"
   PopMsg.msg.Refresh
   
   EsiLogon.lblSqlServer = EsiLogon.txtUsr & " Is Logged On To ES/2000 ERP"
   EsiLogon.lblSqlServer.Refresh
   bUserLoggedOn = 1
'MsgBox "1"
   
End Sub
Public Sub SectionButtonClick(SectionNo As Integer, frm As Form)
   
   'call this on button click events.  example:
   'SectionButtonClick SECTION_
   
   On Error GoTo cmdCl1
   If bSections(SectionNo) = 0 Then
      MsgBox "No User Section Permissions for " & SectionInfo(SectionNo).AppDescription, vbInformation, sSysCaption
      Exit Sub
   End If
   If bUserLoggedOn = 0 Then
      MsgBox "You Must Log On To " & sSysCaption & " ERP", vbInformation, sSysCaption
   Else
      iMinimize = GetSetting("Esi2000", "mngr", "MinOnOpen", iMinimize)
      If SectionInfo(SectionNo).SectionOpen = 0 Then
         SectionInfo(SectionNo).SectionOpen = Shell(sFilePath & SectionInfo(SectionNo).ExeName & " " & Command, vbMaximizedFocus)
      Else
         On Error GoTo cmdCl3
         AppActivate SectionInfo(SectionNo).AppName, True
         SendKeys "% x", False
      End If
      If iMinimize Then frm.WindowState = vbMinimized
   End If
   Exit Sub
   
cmdCl1:
   Resume cmdCl2
cmdCl2:
   ModuleNotFound SectionInfo(SectionNo).AppName
   Exit Sub
   
cmdCl3:
   Resume cmdCl4
cmdCl4:
   SectionInfo(SectionNo).SectionOpen = 0
   
End Sub

Public Sub ShowChoices(WhichButton As Integer, cmd As CommandButton, SectionNo As Integer)
   
   If WhichButton = 2 Then
      On Error Resume Next
      SubjHelp.optFrom = vbUnchecked
      SubjHelp.Section = SectionInfo(SectionNo).AppDescription
      SubjHelp.hlp = SectionInfo(SectionNo).AppButtons
      
      If cmd.Parent.Name = "Esi2000h" Then
         SubjHelp.Left = cmd.Parent.Left + cmd.Left
      Else
         SubjHelp.Top = cmd.Top
         SubjHelp.Left = cmd.Parent.Left + cmd.Left + cmd.Width
      End If
      SubjHelp.Show
      
      ' calculate height
      Dim nCrs As Integer
      Dim nCr As Integer
      Dim nStart As Integer
      nStart = 1
      nCrs = 2 'description line + 1 choice line
      Do While True
         nCr = InStr(nStart, SectionInfo(SectionNo).AppButtons, vbCr)
         If nCr > 0 Then
            nStart = nCr + 1
            nCrs = nCrs + 1
         Else
            Exit Do
         End If
      Loop
      
      SubjHelp.Height = 175 + nCrs * 225 '1500
   End If
   
   
End Sub
