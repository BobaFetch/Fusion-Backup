Attribute VB_Name = "ESIPROJ"
'*** ES/2000 (ES/2001, ES/2002) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'\\ Main Include file for all Projects except ACCRA Tooling
'memory/system

Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Combo Boxes and sample
Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_ADDSTRING = &H143
' SendMessageStr Combo1.hWnd, CB_ADDSTRING, 0&, ByVal sSomeString

'help
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Public Const HELP_CONTEXT = &H1 'Display topic in ulTopic
Public Const HELP_QUIT = &H2 'Terminate help
Public Const HELP_INDEX = &H3 'Display index
Public Const HELP_CONTENTS = &H3
Public Const HELP_HELPONHELP = &H4 'Display help on using help
Public Const HELP_SETINDEX = &H5 'Set the current Index for multi index help
Public Const HELP_SETCONTENTS = &H5
Public Const HELP_CONTEXTPOPUP = &H8
Public Const HELP_FORCEFILE = &H9
Public Const HELP_KEY = &H101 'Display topic for keyword in offabData
Public Const HELP_COMMAND = &H102
Public Const HELP_PARTIALKEY = &H105 'call the search engine in winhelp
Public Const HELP_SHOWTAB = &HB '(11) Show the tab

'ES Color Formats
Public Const ES_RED = &HC0 'jse red for labels
Public Const ES_BLUE = &H800000 'jse blue for labels
Public Const Es_FormBackColor = &H8000000F
Public Const Es_CheckBoxForeColor = &H8000000F
Public Const Es_HelpBackGroundColor = &H80000018
Public Const Es_TextBackColor = &H80000005
Public Const Es_TextForeColor = &H80000008

'About box

Public Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
   dwTotalPhys As Long
   dwAvailPhys As Long
   dwTotalPageFile As Long
   dwAvailPageFile As Long
   dwTotalVirtual As Long
   dwAvailVirtual As Long
End Type

'Window status
Public Const Swp_NOSIZE = &H1
Public Const Swp_NOMOVE = &H2
Public Const hWnd_TopMost = -1
Public Const Hwnd_NOTOPMOST = -2
Public Const Flags = Swp_NOSIZE Or Swp_NOMOVE

'Menu Constants
Public Const MF_BYPOSITION = &H400&
Public Const MF_GRAYED = &H1&
Public Const SC_CLOSE = &HF060
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const SC_MOVE = &HF010
Public Const SC_RESTORE = &HF120

'Cost Constants
Public Const ES_AVERAGECOST As Byte = 0
Public Const ES_STANDARDCOST As Byte = 1

'Form Constants
Public Const ES_RESIZE = 0
Public Const ES_DONTRESIZE = -1
Public Const ES_LIST = 0
Public Const ES_DONTLIST = -1
Public Const ES_IGNOREDASHES = 1 'Compress routine
'MsgBox
Public Const ES_NOQUESTION = &H124 'Question and return (Default NO)
Public Const ES_YESQUESTION = &H24 'Question and return (Default YES)

'StrCase funtion contstants
Public Const ES_FIRSTWORD As Byte = 1

'Cursor types
Public Const ES_FORWARD = 0 'Default
Public Const ES_KEYSET = 1
Public Const ES_DYNAMIC = 2
Public Const ES_STATIC = 3

'RDO Defaults
Public RdoCon As rdoConnection
Public RdoEnv As rdoEnvironment
Public RdoErr As rdoError
Public rdoRes As rdoResultset

'Project variables
Public bAutoCaps As Byte
Public bBold As Byte
Public bCalendar As Byte
Public bEnterAsTab As Byte
Public bSqlRows As Byte
Public bInsertOn As Byte
Public bNextLot As Byte 'Cycle lots
Public bNoCrystal As Boolean
Public bResize As Byte
Public bTestDb As Byte
Public bUserAction As Boolean
Public bvbTest As Byte
Public bVersion As Byte 'Controls Version for changed colums to SQL 7.0

Public iAutoTips As Integer
Public iBarOnTop As Integer
Public iHideRecent As Integer
Public iZoomLevel As Integer

Public lScreenWidth As Long

Public sdsn As String
Public sDataBase As String
Public sFacility As String
Public sFilePath As String
Public sProgName As String
Public sProcName As String
Public sReportPath As String
Public sSaAdmin As String
Public sSaPassword As String
Public sServer As String
Public sSql As String
Public sSysCaption As String

Public sActiveTab(8) As Integer
Public ESI_cmdShowPrint As New EsiKeyBd 'Only one per form
'ini registration
Dim mzBuff As String

'Common Objects

Type CompanyInfo
   Name As String
   Addr(5) As String
   Phone As String
   Fax As String
   GlVerify As Byte
End Type

Public Co As CompanyInfo


Type ModuleErrors
   Number As Long
   Description As String
End Type

Public CurrError As ModuleErrors

Type CurrentSelections
   CurrentPart As String
   CurrentVendor As String
   CurrentCustomer As String
   CurrentShop As String
   CurrentRegion As String
   CurrentGroup As String
   CurrentUser As String
End Type

Public Cur As CurrentSelections


Sub CloseForms()
   On Error GoTo modErr1
   sCurrForm = ""
   bUserAction = True
   Do While Forms.Count > 1
      Unload Forms(1)
   Loop
   DoEvents
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

'Use Windows messaging to fill Combo Strings 8/17/00
'AddComboStr cmbVnd.hWnd, sString

Public Sub AddComboStr(lhWnd As Long, sString As String)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, _
      ByVal "" & Trim(sString)
   
End Sub

'Validate Edits after attemped Update of KeySet
'The operation has falled because a Column in the set
'Was changed somewhere else
'Syntax:   If Err > 0 Then ValidateEdit Me
'Call after .Update command in a Cursor Edit
'2/17/00 cjs

Public Sub ValidateEdit(frm As Form)
   Static bByte As Byte
   
   'Clear it if no error. called from SetDiaPos to reset Static
   'otherwise use example syntax
   If Err = 0 Then bByte = 3
   'Show the error once and again if the continue on
   If bByte = 0 Then
      'This is a 40002 SQL Server subset
      If Left(Err.Description, 5) = "01S03" Then
         MsgBox "The Data Edited Has Been Changed By Another" & vbCr _
            & "Process. You Should Reselect The Data And" & vbCr _
            & "Refresh The Information.", _
            vbInformation, frm.Caption
      Else
         'Process all others
         If Err.Number = 40011 Then
            MsgBox "The Data Has Timed Out. This May Be " _
               & "Normal. Close This Form And Reopen.", _
               vbInformation, frm.Caption
         Else
            If Err.Number <> 40060 Then
               sProcName = "editfunction"
               CurrError.Number = Err.Number
               CurrError.Description = Err.Description
               DoModuleErrors frm
            End If
         End If
      End If
   End If
   bByte = bByte + 1
   If bByte > 3 Then bByte = 0
   Err.Clear
   
End Sub

'Make sure that the user's DSN is pointed to the
'correct server. If none is registered, then build it

Public Function RegisterSqlDsn(sDataSource As String) As String
   Dim sAttribs As String
   If sDataSource = "" Then sDataSource = "ESI2000"
   sAttribs = "Description=" _
              & "ES/2000ERP SQL Server Data " _
              & vbCr & "OemToAnsi=No" _
              & vbCr & "SERVER=" & sServer _
              & vbCr & "Database=" & sDataBase
   'Create new DSN or revise registered DSN.
   rdoEngine.rdoRegisterDataSource sDataSource, _
      "SQL Server", True, sAttribs
   RegisterSqlDsn = sDataSource
   Exit Function
   
modErr1:
   On Error GoTo 0
   RegisterSqlDsn = sDataSource
   
End Function

Public Sub GetFavorites(sSection As String)
   Dim i As Integer
   For i = 1 To 11
      sFavorites(i) = GetSetting("Esi2000", sSection, "Favorite" & Trim(Str(i)), sFavorites(i))
   Next
   sFavorites(i) = GetSetting("Esi2000", sSection, "Favorite" & Trim(Str(i)), sFavorites(i))
   
   For i = 1 To 11
      If sFavorites(i) <> "" Then
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i))).Caption = sFavorites(i)
      End If
   Next
   If sFavorites(i) <> "" Then
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i))).Visible = True
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(Str(i))).Caption = sFavorites(i)
   End If
   
   iBarOnTop = GetSetting("Esi2000", "Programs", "BarOnTop", iBarOnTop)
   iAutoTips = GetSetting("Esi2000", "Programs", "AutoTipsOn", iAutoTips)
   If iAutoTips = 1 Then
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips On"
   Else
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips Off"
   End If
   If iBarOnTop = 1 Then
      MdiSect.SideBar.Visible = False
      MdiSect.TopBar.Visible = True
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
   Else
      MdiSect.SideBar.Visible = True
      MdiSect.TopBar.Visible = False
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
   End If
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", bEnterAsTab)
   sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   If sReportPath = "" Then sReportPath = App.Path & "\"
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   GetCrystalZoom
   
End Sub

Public Sub FillVendors(frm As Form)
   Dim RdoVed As rdoResultset
   On Error GoTo modErr1
   sSql = "Qry_FillVendors"
   bSqlRows = GetDataSet(RdoVed, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         frm.cmbVnd = "" & Trim(!VENICKNAME)
         frm.txtNme = "" & Trim(!VEBNAME)
         frm.lblNme = "" & Trim(!VEBNAME)
         Do Until RdoVed.EOF
            AddComboStr frm.cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoVed = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillvendors"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Sub CloseFiles()
   On Error Resume Next
   Close
   RdoCon.Close
   Set RdoEnv = Nothing
   DoEvents
   
End Sub

Public Sub CheckKeys(KeyCode As Integer)
   'use in KeyDown
   'not for combo boxes or memo fields
   'to use vbKeyinsert you must have a label
   'name InsPanel (or something else)
   bUserAction = True
   If KeyCode = vbKeyDown Then
      SendKeys "{TAB}"
   Else
      If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
   End If
   
End Sub


'Use Windows call to retrieve a specific Help Topic

Sub GetHelpTopic(frm As Form, HelpTopic As Long)
   Dim l&
   l& = WinHelp(frm.hWnd, "esimam.hlp", HELP_KEY, HelpTopic)
   
End Sub




Sub GetCompany(Optional bWantAddress As Byte)
   Dim ActRs As rdoResultset
   Dim bByte As Byte
   Dim a As Integer
   Dim b As Integer
   Dim C As Integer
   Dim d As Integer
   Dim sAddress As String
   
   On Error GoTo modErr1
   If bWantAddress Then
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY,COADR FROM ComnTable " _
             & "WHERE COREF=1"
   Else
      sSql = "SELECT COREF,CONAME,COPHONE,COFAX,COGLVERIFY FROM ComnTable " _
             & "WHERE COREF=1"
   End If
   bSqlRows = GetDataSet(ActRs, ES_STATIC)
   If bSqlRows Then
      With ActRs
         Co.Name = "" & Trim(!CONAME)
         Co.Phone = "" & Trim(!COPHONE)
         Co.Fax = "" & Trim(!COFAX)
         Co.GlVerify = !COGLVERIFY
         If bWantAddress Then sAddress = "" & Trim(!COADR)
      End With
   End If
   'have parse CfLf if we want address for Crystal Reports only
   If bWantAddress Then
      On Error Resume Next
      Err = 0
      a = InStr(1, sAddress, Chr(13) & Chr(10))
      Co.Addr(1) = Left(sAddress, a - 1)
      
      sAddress = Right(sAddress, Len(sAddress) - (a + 1))
      b = InStr(1, sAddress, Chr(13) & Chr(10))
      If b = 0 Then
         bByte = 1
         b = Len(sAddress)
         Co.Addr(2) = Left(sAddress, b)
      Else
         Co.Addr(2) = Left(sAddress, b - 1)
      End If
      
      If bByte = 0 Then
         sAddress = Right(sAddress, Len(sAddress) - (b + 1))
         C = InStr(1, sAddress, Chr(13) & Chr(10))
         If C = 0 Then
            bByte = 1
            C = Len(sAddress)
            Co.Addr(3) = Left(sAddress, C)
         Else
            Co.Addr(3) = Left(sAddress, C - 1)
         End If
      End If
      
      If bByte = 0 Then
         sAddress = Right(sAddress, Len(sAddress) - (C + 1))
         d = InStr(1, sAddress, Chr(13) & Chr(10))
         If d = 0 Then
            bByte = 1
            d = Len(sAddress)
            Co.Addr(4) = Left(sAddress, d)
         Else
            Co.Addr(4) = Left(sAddress, d - 1)
         End If
      End If
   End If
   sFacility = Co.Name
   Set ActRs = Nothing
   Exit Sub
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error GoTo 0
   
End Sub



'Use constants to return cost type
'Public Const ES_AVERAGECOST As Byte = 0
'Public Const ES_STANDARDCOST As Byte = 1

Public Function GetPartCost(sPartRef As String, bCostType As Byte) As Currency
   Dim CostRes As rdoResultset
   
   On Error GoTo modErr1
   sPartRef = Compress(sPartRef)
   sSql = "Qry_GetPartCost '" & sPartRef & "' "
   bSqlRows = GetDataSet(CostRes)
   If bSqlRows Then
      With CostRes
         If bCostType = ES_AVERAGECOST Then
            GetPartCost = !PAAVGCOST
         Else
            GetPartCost = !PASTDCOST
         End If
         .Cancel
      End With
   End If
   Set CostRes = Nothing
   Exit Function
   
modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Moderr2
Moderr2:
   On Error GoTo 0
   GetPartCost = 0
   
End Function

Sub GridKeyCheck(KeyAscii As Integer)
   'Key trap for grids
   'If KeyAscii = 13 Then
   '    KeyAscii = 0
   '    Exit Sub
   'End If
   If KeyAscii > 32 Then SendKeys "+{RIGHT}{DEL}"
   
End Sub

Sub GridKeyDate(KeyAscii As Integer)
   'Date field used in Keypress Event for Grids
   'Allows certain characters to be honored
   
   If KeyAscii = 13 Then
      KeyAscii = 0
      Exit Sub
   End If
   If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
   If KeyAscii = 8 Then Exit Sub
   Select Case KeyAscii
      Case Is < 43
         KeyAscii = 0
      Case Is > 57
         KeyAscii = 0
      Case Is = 45, 46
         KeyAscii = 47
   End Select
   
End Sub

Sub GridKeyValue(KeyAscii As Integer)
   'Numeric field used in Keypress Event for Grids
   'Allows certain characters to be honored
   
   If KeyAscii = 13 Then
      KeyAscii = 0
      Exit Sub
   End If
   If KeyAscii = 32 Then Exit Sub
   If KeyAscii = 8 Then Exit Sub
   If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
   Select Case KeyAscii
      Case Is < 43
         KeyAscii = 0
      Case Is > 57
         KeyAscii = 0
      Case Is = 47
         KeyAscii = 0
   End Select
   
End Sub

Sub KeyCase(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      On Error Resume Next
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
      If Not bInsertOn Then
         If Len(MdiSect.ActiveForm.ActiveControl) > 0 Then _
                If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
      End If
   End If
   
End Sub

Public Sub KeyCheck(KeyAscii As Integer)
   'Check key for Enter key
   'syntax in Keypress Procedure KeyCheck KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      On Error Resume Next
      If Not bInsertOn Then
         If Len(MdiSect.ActiveForm.ActiveControl) > 0 Then _
                If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
      End If
   End If
   
End Sub

Sub KeyDate(KeyAscii As Integer)
   'Changes ".", " " and "-" to "/" for dates
   'syntax in Keypress: KeyDate KeyAscii
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      If KeyAscii = 8 Or KeyAscii = 9 Then Exit Sub
      Select Case KeyAscii
         Case Is < 43
            KeyAscii = 0
         Case Is > 57
            KeyAscii = 0
         Case Is = 45, 46
            KeyAscii = 47
      End Select
      If Not bInsertOn Then
         If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
      End If
   End If
   
End Sub

Sub KeyTime(KeyAscii)
   'Time field used in Keypress Event
   'Allows certain characters to be honored
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
         Exit Sub
      End If
   Else
      If Not bInsertOn Then If KeyAscii > 13 Then SendKeys "{DEL}"
      If KeyAscii = 8 Or KeyAscii = 9 Then Exit Sub
      If KeyAscii = 65 Then KeyAscii = 97
      If KeyAscii = 80 Then KeyAscii = 112
      If KeyAscii = 45 Then KeyAscii = 58
      If KeyAscii = 46 Then KeyAscii = 58
      If KeyAscii = 58 Or KeyAscii = 97 Or KeyAscii = 112 Then Exit Sub
      Select Case KeyAscii
         Case 43
            KeyAscii = 112
         Case Is < 43, 47, Is > 57
            KeyAscii = 0
      End Select
   End If
   
End Sub

Sub KeyValue(KeyAscii)
   'Allows only numbers, "-" and "." for value
   'fields like money or quantities
   'syntax in Keypress: KeyValue KeyAscii
   bUserAction = True
   If KeyAscii = 13 Then
      If bEnterAsTab Then
         KeyAscii = 0
         SendKeys "{TAB}"
      End If
   Else
      Select Case KeyAscii
         Case 8, 32, 43 To 46, 48 To 57
         Case Else
            KeyAscii = 0
      End Select
      If Not bInsertOn Then
         If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
      End If
   End If
   
End Sub


Sub MouseCursor(MCursor As Integer)
   'Allows consistant MousePointer Updates
   Screen.MousePointer = MCursor
   bUserAction = True
   
End Sub


'Server and DSN Registery moved to GetRecentList() 6/12/99
'Provisions for TestDB 3/7/02

Sub OpenSqlServer(Optional bReStart As Boolean)
   Dim b As Byte
   Dim sConnect As String
   MouseCursor 11
   Dim iFreeFile As Integer
   On Error GoTo PrjOs1
   sSaAdmin = Trim(GetSysLogon(True))
   sSaPassword = Trim(GetSysLogon(False))
   sSysCaption = GetSystemCaption()
   
   'Use upper case here
   'sServer = "AWI_SQL_SVR": sSaAdmin = "sa": sSaPassword = "":  bvbTest = 1    'Austin Waterjet 6.5
   'sServer = "ESI_SQL_SVR": sSaAdmin = "sa": sSaPassword = "": bvbTest = 1     'Company
   'sServer = "ironhorseserver": sSaAdmin = "sa": sSaPassword = "": bvbTest = 1 'Iron Horse 6.5
   'sServer = "JEVCO2": sSaAdmin = "sa": sSaPassword = "": bvbTest = 1          'SQL Server 7.0
   
   '**** Need to unblock these for testing customers ****
   'sSaAdmin = "sa"
   'sSaPassword = ""
   
   sServer = UCase$(sServer)
   GetCurrentDatabase
   'Check and reset for TestDb
   If Trim(sFilePath) = "" Then
      If sServer = "ESI_DEV_SVR" Then
         sFilePath = "c:\esi2000\"
         bvbTest = 1
      Else
         sFilePath = App.Path & "\"
      End If
   End If
   Set RdoEnv = rdoEnvironments(0)
   RdoEnv.CursorDriver = rdUseIfNeeded
   Set RdoCon = RdoEnv.OpenConnection(dsName
   = "", _
     Prompt
   = rdDriverNoPrompt, _
     Connect
   = "uid=" & sSaAdmin & ";pwd=" & sSaPassword & ";driver={SQL Server};" _
     & "server=" & sServer & ";database=" & sDataBase & ";")
   
   '**** Check SQL Server Version here ****
   sSql = "SELECT COADR FROM ComnTable"
   bSqlRows = GetDataSet(rdoRes, ES_FORWARD)
   If bSqlRows Then
      'Column was changed in 7.0 from Text (BLOB) to varchar 255
      'To handle larger varchar columns
      If rdoRes.rdoColumns(0).Size > 255 Then bVersion = 6 Else bVersion = 7
      rdoRes.Cancel
   End If
   'for testing after SQL Server is open
   'sServer = "AWI_SQL_SVR"
   'sServer = "IRONHORSESERVER"
   'sServer = "JEVCO2"
   If Not bReStart Then
      UpdateTables
      b = CheckSecuritySettings()
      'For now all are 0
      'b = 0
      If b = 0 Then GetSectionPermissions
      GetCompany
   Else
      MouseCursor 0
   End If
   Exit Sub
   
PrjOs1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume PrjOs2
PrjOs2:
   MouseCursor 0
   On Error GoTo 0
   MsgBox LTrim(Str(CurrError.Number)) & vbCr & CurrError.Description & vbCr _
                & "Unable To Make SQL Server Connection.", 48, "ESI2002 ERP"
End

End Sub

Sub ParseComment(TestCmt As Variant)
   'Replace double quotes for SQL Server text fields
   Dim a As Integer
   Dim d As Integer
   Dim E As Integer
   Dim g As Integer
   Dim k As Integer
   Dim n As Integer
   Dim NewComment As String
   
   On Error GoTo modErr1
   NewComment = RTrim(TestCmt)
   a = Len(NewComment)
   NewComment = NewComment & Chr$(255)
   k = 1
   Do Until k > a
      If Mid(NewComment, k, 1) = Chr(34) Then
         Mid(NewComment, k, 1) = Chr(39)
         NewComment = NewComment & Chr$(255)
         E = Len(NewComment)
         For g = E To k + 1 Step -1
            Mid(NewComment, g, 1) = Mid(NewComment, g - 1, 1)
         Next
         Mid(NewComment, k + 1, 1) = Chr(39)
         k = k - 1
      End If
      k = k + 1
   Loop
   
   NewComment = RTrim(NewComment)
   a = Len(NewComment)
   k = a
   For n = k To 1 Step -1
      If Mid(NewComment, n, 1) = Chr(255) Then a = a - 1
   Next
   NewComment = Left(NewComment, a)
   TestCmt = NewComment
   Exit Sub
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error GoTo 0
   
End Sub





'Used to set an MDIChild form position and provide ToolTips
'See ES_LIST, ES_DONTLIST Constants
'See ES_RESIZE, ES_DONTRESIZE Constansts

Sub SetDiaPos(frm As Form, Optional DontList As Boolean, Optional noResize As Boolean)
   Dim i As Integer
   If Not noResize Then SetFormSize frm
   frm.Move 0, 0
   bUserAction = True
   bNextLot = 0
   Err.Clear
   ValidateEdit frm
   On Error Resume Next
   sProcName = ""
   frm.cmdCan.Cancel = True
   frm.KeyPreview = True
   If Not DontList Then i = SetRecent(frm)
   GetCurrentSelections
   
   'Setup the controls and bring some order to the place. Like it or not
   'colors 10/2/01
   frm.BackColor = Es_FormBackColor
   For i = 0 To frm.Controls.Count - 1
      If TypeOf frm.Controls(i) Is TextBox Then
         frm.Controls(i).BackColor = Es_TextBackColor
         frm.Controls(i).ForeColor = Es_TextForeColor
         If frm.Controls(i).Tag <> 5 Then
            frm.Controls(i).Text = " "
            If Trim(frm.Controls(i).ToolTipText) = "" Then
               If frm.Controls(i).Tag = 1 Then
                  frm.Controls(i).ToolTipText = "Value Formatted TextBox"
               ElseIf frm.Controls(i).Tag = 2 Then
                  frm.Controls(i).ToolTipText = "Any Format TextBox"
               ElseIf frm.Controls(i).Tag = 3 Then
                  frm.Controls(i).ToolTipText = "UpperCase Formatted TextBox"
               ElseIf frm.Controls(i).Tag = 4 Then
                  frm.Controls(i).ToolTipText = "Date Formatted TextBox"
               ElseIf frm.Controls(i).Tag = 5 Then
                  frm.Controls(i).ToolTipText = "Time Formatted TextBox"
               End If
            End If
         Else
            frm.Controls(i).Text = "  :  "
         End If
      ElseIf TypeOf frm.Controls(i) Is CommandButton Then
         frm.Controls(i).BackColor = Es_FormBackColor
         frm.Controls(i).ForeColor = Es_TextForeColor
      ElseIf TypeOf frm.Controls(i) Is SSRibbon Then
         frm.Controls(i).BackColor = Es_FormBackColor
      ElseIf TypeOf frm.Controls(i) Is SSFrame Or TypeOf frm.Controls(i) Is Frame Then
         frm.Controls(i).BackColor = Es_FormBackColor
      ElseIf TypeOf frm.Controls(i) Is SSPanel Then
         frm.Controls(i).BackColor = Es_FormBackColor
      ElseIf TypeOf frm.Controls(i) Is CheckBox Then
         frm.Controls(i).BackColor = Es_FormBackColor
         If Left$(frm.Controls(i).Caption, 2) = "__" Then
            frm.Controls(i).ForeColor = Es_CheckBoxForeColor
         Else
            frm.Controls(i).ForeColor = Es_TextForeColor
         End If
         If Trim(frm.Controls(i).ToolTipText) = "" Then _
                 frm.Controls(i).ToolTipText = "CheckBox SpaceBar Or Click To Select"
         
      ElseIf TypeOf frm.Controls(i) Is Frame Or TypeOf frm.Controls(i) Is OptionButton Then
         frm.Controls(i).BackColor = Es_FormBackColor
         frm.Controls(i).ForeColor = Es_TextForeColor
      Else
         If TypeOf frm.Controls(i) Is ComboBox Then
            frm.Controls(i).BackColor = Es_TextBackColor
            frm.Controls(i).ForeColor = Es_TextForeColor
            If frm.Controls(i).Tag <> "" Then
               If frm.Controls(i).Tag = 4 Then
                  frm.Controls(i).ToolTipText = "Date As 09/01/02, " _
                               & "09.01.02, 090102 Or Pull Down"
               End If
               If Trim(frm.Controls(i).ToolTipText) = "" Then
                  If frm.Controls(i).Tag = 1 Then
                     frm.Controls(i).ToolTipText = "Value Formatted ComboBox"
                  ElseIf frm.Controls(i).Tag = 2 Then
                     frm.Controls(i).ToolTipText = "Any Format ComboBox"
                  ElseIf frm.Controls(i).Tag = 3 Then
                     frm.Controls(i).ToolTipText = "UpperCase Formatted ComboBox"
                  ElseIf frm.Controls(i).Tag = 8 Then
                     frm.Controls(i).ToolTipText = "Locked Edit ComboBox"
                  End If
               End If
            End If
         End If
      End If
   Next
   'Tool tips
   If iAutoTips = 1 Then
      For i = 0 To frm.Controls.Count - 1
         If TypeOf frm.Controls(i) Is TextBox Then
            If frm.Controls(i).ToolTipText = "" Then _
                            frm.Controls(i).Text = " "
            If frm.Controls(i).ToolTipText = "" Then
               Select Case Val(frm.Controls(i).Tag)
                  Case 1
                     frm.Controls(i).ToolTipText = "Value (number)"
                  Case 3
                     frm.Controls(i).ToolTipText = "Upper Case Entry"
                  Case 4
                     frm.Controls(i).ToolTipText = "Date as 09/15/99,09-15-99 or 091599"
                  Case 5
                     frm.Controls(i).ToolTipText = "Time as 10:32"
                     frm.Controls(i).Text = "  :  "
                  Case 9
                     frm.Controls(i).ToolTipText = "Multiple Line Entry"
                  Case Else
                     frm.Controls(i).ToolTipText = "Any Alpa/Numeric Entry"
               End Select
            End If
         Else
            If TypeOf frm.Controls(i) Is ComboBox Then
               If frm.Controls(i).ToolTipText = "" Or _
                                 frm.Controls(i).Tag <> 4 Then _
                                 frm.Controls(i).ToolTipText = "ComboBox"
               End If
            End If
            If TypeOf frm.Controls(i) Is CommandButton Then
               If frm.Controls(i).Name = "cmdOk" Then _
                               frm.Controls(i).ToolTipText = "Continue Processing"
               If frm.Controls(i).Name = "cmdCan" Then _
                               frm.Controls(i).ToolTipText = "Close Form (Escape)"
               If frm.Controls(i).ToolTipText = "" Then _
                               frm.Controls(i).ToolTipText = "Command Button"
            Else
               If TypeOf frm.Controls(i) Is ListBox Then
                  If frm.Controls(i).ToolTipText = "" Then _
                                  frm.Controls(i).ToolTipText = "ListBox"
               End If
            End If
            If TypeOf frm.Controls(i) Is SSRibbon Then
               If frm.Controls(i).Name = "optDis" Then
                  frm.Controls(i).ToolTipText = "Display The Report"
               Else
                  If frm.Controls(i).Name = "optPrn" Then _
                                  frm.Controls(i).ToolTipText = "Print The Report"
               End If
            End If
         Next
      End If
      
   End Sub
   
   
   
   'Small message instead of a MsgBox
   'Timer On either True or False
   'Syntax: SysMsg "User Message", True
   'SysMessage may be up to 24 characters

   Sub Sysmsg(SysMessage As String, TimerOn As Byte, Optional frm As Form)
      On Error GoTo modErr1
      PopMsg.tmr1.Enabled = TimerOn
      PopMsg.msg = SysMessage
      PopMsg.Show vbModal
      On Error Resume Next
      frm.Refresh
      Exit Sub
      
   modErr1:
      Resume Moderr2
   Moderr2:
      'Can't show modal form on MdiChildren
      On Error Resume Next
      PopMsg.tmr1.Enabled = TimerOn
      PopMsg.msg = SysMessage
      PopMsg.Show
      
   End Sub
   
   'Test for a valid date otherwise Use Today
   'Syntax:  txtDte = CheckDate(txtDte)

   Function CheckDate(NewDate As String)
      Dim a As Integer
      Dim l As Long
      
      On Error GoTo modErr1
      NewDate = Trim(NewDate)
      If Len(NewDate) > 8 Then
         NewDate = Left(NewDate, 8)
      Else
         If Len(NewDate) = 0 Then NewDate = Format(Now, "mm/dd/yy")
      End If
      a = InStr(1, NewDate, "/")
      If a = 0 Then
         NewDate = Format(NewDate, "00/00/00")
      Else
         NewDate = Format(NewDate, "mm/dd/yy")
      End If
      If Val(Left(NewDate, 2)) < 10 Then
         If Mid(NewDate, 2, 1) = "/" Then NewDate = "0" & NewDate
      End If
      If Len(NewDate) = 6 Then
         If Val(Mid(NewDate, 1, 2)) > 0 And Val(Mid(NewDate, 3, 2)) > 0 And Val(Mid(NewDate, 5, 2)) > 0 Then
            NewDate = Left(NewDate, 2) & "/" & Mid(NewDate, 3, 2) & "/" & Right(NewDate, 2)
         End If
      End If
      If Len(NewDate) < 4 Then NewDate = NewDate & Right(Format(Now, "mm/dd/yy"), 2)
      NewDate = Format(NewDate, "mm/dd/yy")
      l& = DateValue(NewDate)
      CheckDate = NewDate
      Exit Function
      
   modErr1:
      On Error Resume Next
      Beep
      CheckDate = Format(Now, "mm/dd/yy")
      
   End Function


   Public Sub KeyMemo(KeyAscii As Integer)
      If Not bInsertOn Then
         If KeyAscii > 32 Then SendKeys "+{RIGHT}{DEL}", True
      End If
      
   End Sub
   
   Public Sub SetMdiReportsize(frm As Form)
      'Sets report size based on monitor size
      'requires as large as possible for Win9x generic
      'monitor
      Dim bWindowSize As Byte
      Dim a As Integer
      Dim b As Integer
      
      On Error Resume Next
      MdiSect.crw.Reset
      GetCrystalConnect
      sSql = ""
      'clear any report variables
      'Resolve a bug in Crystal that doesn't clear a report
      For b = 0 To 60
         MdiSect.crw.Formulas(b) = ""
         MdiSect.crw.SectionFormat(b) = ""
         MdiSect.crw.SectionFont(b) = ""
      Next
      a = Screen.TwipsPerPixelX
      b = Screen.TwipsPerPixelY
      bUserAction = True
      bWindowSize = GetSetting("Esi2000", "System", "ReportMax", bWindowSize)
      If bWindowSize = 0 Then
         If iBarOnTop = False Then
            MdiSect.crw.WindowState = 0
            MdiSect.crw.WindowTop = 650 / b
            MdiSect.crw.WindowHeight = (MdiSect.Height / b) - (1100 / b)
            MdiSect.crw.WindowLeft = 1960 / a
            MdiSect.crw.WindowWidth = (MdiSect.Width / a) - (2120 / a)
         Else
            MdiSect.crw.WindowTop = 1250 / b
            MdiSect.crw.WindowHeight = (MdiSect.Height / b) - (1700 / b)
            MdiSect.crw.WindowLeft = 600 / a
            MdiSect.crw.WindowWidth = (MdiSect.Width / a) - (750 / a)
         End If
      Else
         MdiSect.crw.WindowState = 2
         MdiSect.crw.WindowTop = 0
         MdiSect.crw.WindowHeight = Screen.Height
         MdiSect.crw.WindowLeft = 0
         MdiSect.crw.WindowWidth = Screen.Width
      End If
      On Error GoTo 0
      
   End Sub
   
   'Adjust Box length to fit data fields
   'Also checks to make sure there and no (') to mess SQL Server up

   Public Function CheckLen(sTextBox As String, iTextLength As Integer) As String
      sTextBox = Trim(sTextBox)
      If Len(sTextBox) > iTextLength Then
         Beep
         sTextBox = Left(sTextBox, iTextLength)
      End If
      CheckLen = sTextBox
      iTextLength = InStr(1, CheckLen, Chr$(39))
      If iTextLength > 0 Then CheckLen = CheckComments(CheckLen)
      
   End Function
   
   
   'Check Time entry in a TextBox time formated
   'syntax  is txtTme = GetTime (txtTme)

   Public Function GetTime(TimeEntry As Variant) As Variant
      Dim i As Integer
      On Error GoTo modErr1
      i = Len(Trim(TimeEntry))
      Select Case i
         Case 1
            If Val(TimeEntry) > 0 Then TimeEntry = "0" & TimeEntry & ":00"
         Case 2
            If Right(TimeEntry, 1) = ":" Then
               TimeEntry = "0" & TimeEntry & "00"
            Else
               If Val(TimeEntry) > 0 Then TimeEntry = TimeEntry & ":00"
            End If
         Case 3
            If Right(TimeEntry, 1) = ":" Then
               TimeEntry = TimeEntry & "00"
            Else
               If Val(Left(TimeEntry, 2)) > 0 Then TimeEntry = TimeEntry & "00"
            End If
         Case 4
            If Mid(TimeEntry, 3, 1) = ":" Then TimeEntry = TimeEntry & "0"
      End Select
      TimeEntry = TimeValue(TimeEntry)
      GetTime = Format(TimeEntry, "hh:nna/p")
      Exit Function
      
   modErr1:
      GetTime = ""
      On Error GoTo 0
      
   End Function
   
   
   '11/15/00 changed the notification type (confusion)...see end
   '
   '8/19/99 added sProcName to the log
   '
   'With V5(SP2) the following can be used in all
   'Procedures.  Need not be unique
   '
   '   On Error Goto DiaErr1
   '       code....
   '   Exit Sub, Function
   '
   'DiaErr1:
   '   sProcName = "fillcombo"
   '   CurrError.Number = Err.Number
   '   CurrError.Description = Err.Description
   '   DoModuleErrors me
   '
   'End sub
   'Added EsiError.log 6/9/99

   Public Sub DoModuleErrors(frm As Form)
      Dim bByte As Byte
      Dim iFreeFile As Integer
      Dim iWarningType As Integer
      Dim smsg As String
      Dim sMsg2 As String
      
      'error log
      Dim sDate As String * 16
      Dim sSection As String * 8
      Dim sForm As String * 12
      Dim sErrNum As String * 10
      Dim sErrSev As String * 2
      Dim sUserName As String * 20
      
      MouseCursor 13
      On Error Resume Next
      If UCase$(sServer) = "ESI_DEV_SVR" Then
         sFilePath = "c:\esi2000\"
      Else
         sFilePath = App.Path & "\"
      End If
      '   sFilePath = "c:\esi2000\"
      iFreeFile = FreeFile
      Open sFilePath & "EsiError.log" For Append Shared As iFreeFile
      
      sDate = Format(Now, "mm/dd/yy hh:mm")
      sSection = Left(sProgName, 5)
      sForm = frm.Name
      sErrNum = Str(CurrError.Number)
      sUserName = Left(Cur.CurrentUser, 18)
      
      'Defualt Warning Flag. Setting bByte to True changes
      'the Warning Flag and smooth closes the app if req'd.
      iWarningType = vbExclamation
      
      Select Case CurrError.Number
         Case 3
            smsg = "Return Without GoSub"
            bByte = 0
         Case 5
            smsg = "Invalid Procedure Call"
            bByte = 0
         Case 6
            smsg = "Overflow"
            bByte = 0
         Case 7
            smsg = "Out Of Memory"
            bByte = 1
         Case 9
            smsg = "Subscript Out Of Range"
            bByte = 0
         Case 10
            smsg = "This Array Is Fixed Or Temporarily Locked"
            bByte = 1
         Case 11
            smsg = "Division By Zero"
            bByte = 1
         Case 13
            smsg = "Type Mismatch"
            bByte = 0
         Case 14
            smsg = "Out Of String Space"
            bByte = 0
         Case 16
            smsg = "Expression Too Complex"
            bByte = 0
         Case 17
            smsg = "Can't Perform Requested Operation"
            bByte = 0
         Case 18
            smsg = "User Interrupt Occurred"
            bByte = 1
         Case 20
            smsg = "Resume Without Error"
            bByte = 0
         Case 28
            smsg = "Out Of Strack Space"
            bByte = 1
         Case 35
            smsg = "Sub, Function, Or Property Not Defined"
            bByte = 0
         Case 47
            smsg = "Too Many DLL Application Clients"
            bByte = 1
         Case 48
            smsg = "Error In Loading DLL"
            bByte = 1
         Case 49
            smsg = "Bad DLL Calling Convention"
            bByte = 1
         Case 51
            smsg = "Internal Error"
            bByte = 1
         Case 52
            smsg = "Bad File Name Or Number"
            bByte = 1
         Case 53
            smsg = "File Not Found"
            bByte = 0
            iWarningType = vbInformation
         Case 54
            smsg = "Bad File Mode"
            bByte = 0
         Case 55
            smsg = "File Already Open"
            bByte = 0
         Case 57
            smsg = "Device I/O Error"
            bByte = 0
         Case 58
            smsg = "File Already Exists"
            bByte = 0
         Case 59
            smsg = "Bad Record Length"
            bByte = 1
         Case 61
            smsg = "Disk Full"
            bByte = 1
         Case 62
            smsg = "Input Past End Of File"
            bByte = 0
         Case 63
            smsg = "Bad Record Number"
            bByte = 0
         Case 67
            smsg = "Too Many Files"
            bByte = 1
         Case 68
            smsg = "Device Unavailable"
            bByte = 0
         Case 70
            smsg = "Permission Denied"
            bByte = 0
         Case 71
            smsg = "Disk Not Ready"
            bByte = 0
         Case 74
            smsg = "Can't Rename With Different Drive"
            bByte = 0
         Case 75
            smsg = "Path/File Access Error"
            bByte = 0
         Case 76
            smsg = "Path Not Found"
            bByte = 0
         Case 91
            smsg = "Object Variable Or With Block Variable Not Set."
            smsg = smsg & vbCr & "Check Network Connection"
            bByte = 0
         Case 94
            smsg = "Invalid Use Of Null. Please Report This Error."
            bByte = 0
         Case 380
            smsg = "Invalid Property Value"
            bByte = 0
         Case 438
            smsg = "Function Not Supported By Object.  Please Report This Error."
            bByte = 0
         Case 482, 483, 486
            smsg = "Printer Error"
            bByte = 0
            'Jet
         Case 3001 To 3648
            smsg = "JET DSS Database Error. Contact Systems Administrator."
            bByte = 0
            'Crystal
         Case 20500
            smsg = "Not Enough Memory To Complete Report. " & vbCr _
                   & "SSCSDK32.DLL May Be Missing Or Corrupt."
            bByte = 0
         Case 20501 To 20506
            smsg = "Crystal Reports Documentation. " & vbCr _
                   & "Contact Systems Administrator."
            bByte = 0
         Case 20507
            smsg = "Report Wasn't Found Or Couldn't Be Loaded. " & vbCr _
                   & "Check Your Report Path In Settings."
            bByte = 0
         Case 20508 To 20514
            smsg = "Crystal Reports Documentation. " & vbCr _
                   & "Contact Systems Administrator."
            bByte = 0
         Case 20515
            smsg = "Error In Selection Formula. " & vbCr _
                   & "Contact Systems Administrator And ESI."
            bByte = 0
         Case 20516, 20517
            smsg = "Not Windows Resources To Complete Report. " & vbCr _
                   & "Close Some Applications,  " & sSysCaption & " And Restart."
            bByte = 0
         Case 20518
            smsg = "The Report Section Formatted Does Not Exist. " & vbCr _
                   & "Please Report This Warning And The Report Name."
            bByte = 0
         Case 20519
            smsg = "Not Windows Resources To Complete Report. " & vbCr _
                   & "Close Some Applications, " & sSysCaption & " And Restart."
            bByte = 0
         Case 20520
            smsg = "Print Job Started And Report In Progress, " & vbCr _
                   & "There Is No Default Print Or The Printer Is Offline." & vbCr _
                   & "Crystal Reports Notice Not An Error."
            iWarningType = vbInformation
            bByte = 0
         Case 20521 To 20522
            smsg = "Not Windows Resources To Complete Report. " & vbCr _
                   & "Close Some Applications, " & sSysCaption & " And Restart."
            bByte = 0
         Case 20523, 20524
            smsg = "Crystal Reports Documentation. " & vbCr _
                   & "Contact Your Systems Administrator."
            bByte = 0
         Case 20525
            smsg = "Report Is Damaged. Unable To Open Report. " & vbCr _
                   & "Contact Your Systems Administrator."
            bByte = 0
         Case 20526
            smsg = "No Default Printer Has Been Set. "
            bByte = 0
         Case 20527
            smsg = "Error In SQL Server Connection. " & vbCr _
                   & "Check Crystal Report Settings Or Query."
            bByte = 0
         Case 20529
            smsg = "Your Disk Drive Is Full And Files May Be Lost." & vbCr _
                   & "Exit " & sSysCaption & " And Free Resources."
            bByte = 0
         Case 20530
            smsg = "File I/O Error. Disk Problem Other Than Full." & vbCr _
                   & "Exit " & sSysCaption & " Contact Systems Administrator."
            bByte = 0
         Case 20531
            smsg = "Incorrect Password. Permission Denied."
            bByte = 0
         Case 20532
            smsg = "File I/O Error. Disk Problem Other Than Full." & vbCr _
                   & "Exit " & sSysCaption & " Contact Systems Administrator."
            bByte = 0
         Case 20533
            smsg = "Unable To Open The Database File." & vbCr _
                   & "Contact Your Systems Administrator."
            bByte = 0
         Case 20534
            smsg = "Database DLL Error Or The Database Is In Use." & vbCr _
                   & "Contact Your Systems Administrator."
            bByte = 0
         Case 20535 To 20543
            smsg = CurrError.Description
            bByte = 0
         Case 20544
            smsg = "This Report Is Open By Another User." & vbCr _
                   & "Try The Report Again In A Few Minutes."
            bByte = 0
         Case 20545
            smsg = CurrError.Description
            iWarningType = vbInformation
            bByte = 0
         Case 20546 To 20598
            smsg = CurrError.Description
            bByte = 0
         Case 20599
            smsg = "ODBC Permissions/Access Error. Check ODBC Data Source." & vbCr _
                   & "DSN " & sdsn & " May Be Improperly Installed Or Does Not Exist."
            bByte = 0
         Case 20600 To 20996
            smsg = "Undocumented Crystal Reports Error." & vbCr _
                   & "Contact Your Systems Administrator."
         Case 20997
            smsg = "Invalid Report Path Or No Network Permissions." & vbCr _
                   & "Check Your Report Path And Server Permissions."
            bByte = 0
         Case 20998
            smsg = "Report Path Is Too Long.     " & vbCr _
                   & "Use A Mapped Path Name Instead (x:\somedir) Or " _
                   & "Possible Mismatch Of Graph DLL Libraries."
            bByte = 0
            'Rdo
         Case 40000
            smsg = "An Error Occurred Configuring The DataSource Name."
            bByte = 0
         Case 40001
            smsg = "SQL Returned No Data Found From Query."
            iWarningType = vbInformation
            bByte = 0
         Case 40002
            If Left(CurrError.Description, 5) = "01000" And Left(CurrError.Description, 5) = "08S01" Then
               smsg = Left(CurrError.Description, 5) & "-Attempted To Enter An Illegal Character " & vbCr _
                      & "Or The Entry Is Too Long."
               bByte = 0
            Else
               If Left(CurrError.Description, 5) = "S0002" Then
                  smsg = Left(CurrError.Description, 5) & "-The Requested Table Wasn't Found."
                  bByte = 0
               Else
                  If Left(CurrError.Description, 5) = "S0022" Then
                     smsg = Left(CurrError.Description, 5) & "-The Requested Column Wasn't Found."
                     bByte = 0
                  Else
                     If Left(CurrError.Description, 5) = "01000" Then
                        smsg = "An Attempt Was Made To Add A Duplicate Record."
                        bByte = 0
                     Else
                        If InStr(CurrError.Description, "0851") > 0 Then
                           smsg = "ODBC Link Was Lost. Reconnection Required."
                           bByte = 1
                        Else
                           smsg = "Internal ODBC Error Encountered."
                           bByte = 0
                        End If
                     End If
                  End If
               End If
               If Left(CurrError.Description, 5) = "37000" Then
                  'Changed some for SSL 7.0
                  smsg = Left(CurrError.Description, 5) & vbCr _
                         & "The Cursor Is No Longer Open. Invalid Character " & vbCr _
                         & "Found Or The Database Transaction Log Is Full. " & vbCr _
                         & "Please Report This To Your Systems Administrator."
                  bByte = 0
               End If
            End If
         Case 40003
            smsg = "An Invalid Value For The Cursor Driver Was Passed."
            bByte = 0
         Case 40004
            smsg = "An Invalid ODBC Handle Was Encountered."
            bByte = 0
         Case 40005
            smsg = "Invalid Connection String."
            bByte = 1
         Case 40006
            smsg = "An Unexpected Error Occurred."
            bByte = 0
         Case 40008
            smsg = "Invalid Operation For Forward-Only Cursor."
            bByte = 0
         Case 40009
            smsg = "No Current Row (No Matching Query Data Found)."
            iWarningType = vbInformation
            bByte = 0
         Case 40010
            smsg = "Invalid Row For Add New."
            bByte = 0
         Case 40011
            smsg = "Object Is Invalid Or Not Set."
            bByte = 0
         Case 40012
            smsg = "Invalid Seek Flag."
            bByte = 0
         Case 40013
            smsg = "Partial Equality Requires String Column."
            bByte = 0
         Case 40014
            smsg = "Incompatible Data Types For Compare."
            bByte = 0
         Case 40015
            smsg = "Can't Create Prepared Statement."
            bByte = 0
         Case 40016
            smsg = "Version.DLL Error."
            bByte = 1
         Case 40017, 40018
            smsg = "Can't Execute Statement."
            bByte = 0
         Case 40019
            smsg = "An Invalid Value For The Concurrency Option."
            bByte = 0
         Case 40020
            smsg = "Can't Open Result Set For Unnamed Table."
            bByte = 0
         Case 40021
            smsg = "Object Collection Error."
            bByte = 0
         Case 40022
            smsg = "The RDO Results Set Is Empty (No Data)."
            iWarningType = vbInformation
            bByte = 0
         Case 40023
            smsg = "Invalid State For Cursor Move. "
            bByte = 0
         Case 40024
            smsg = "Already Beyond The End Of The Result Set."
            bByte = 0
         Case 40025
            smsg = "BOF Already Set."
            bByte = 0
         Case 40026
            smsg = "Invalid Result Set State For Update."
            bByte = 0
         Case 40027
            smsg = "Invalid Bookmark Or No Bookmark Allowed."
            bByte = 0
         Case 40028
            smsg = "Invalid Bookmark Argument To Move."
            bByte = 0
         Case 40029
            smsg = "Current Row As EOF/BOF Already Set."
            bByte = 0
         Case 40030
            smsg = "Already At BOF."
            bByte = 0
         Case 40031
            smsg = "Already At EOF."
            bByte = 0
         Case 40032
            smsg = "Couldn't Load The ODBC Installation Library."
            bByte = 1
         Case 40033
            smsg = "An Invalid Value For The Prompt Option Was Passed."
            bByte = 1
         Case 40034
            smsg = "An Invalid Value For The Cursor Type Parameter Was Passed."
            bByte = 0
         Case 40035
            smsg = "Column Not Bound Correctly."
            bByte = 0
         Case 40036
            smsg = "Unbound Column-Use Get Chunk Method."
            bByte = 0
         Case 40037
            smsg = "Can't Assign Value To Unbound Column."
            bByte = 0
         Case 40038
            smsg = "Can't Assign Value To Non-Updatable Field."
            bByte = 0
         Case 40039
            smsg = "Can't Assign Value To Column Unless In Edit Mode."
            bByte = 0
         Case 40040
            smsg = "Incorrect Type For Parameter."
            bByte = 0
         Case 40041
            smsg = "Object Collection: Couldn't Find Column Requested By Query."
            bByte = 0
         Case 40042
            smsg = "Can't Assign Value To Unbound Parameter."
            bByte = 0
         Case 40043
            smsg = "Can't Assign Value To Output-Only Parameter."
            bByte = 0
         Case 40044
            smsg = "Incorrect RDO Parameter Type."
            bByte = 0
         Case 40045
            smsg = "Tried To Execute A Query With An Asynchronous Query In Progress."
            bByte = 0
         Case 40046
            smsg = "The Object Has Already Been Closed."
            bByte = 0
         Case 40047
            smsg = "Invalid Name For The Environment."
            bByte = 0
         Case 40048
            smsg = "Environment Name Already Exists In The Collection."
            bByte = 0
         Case 40049
            smsg = "Object Collection Is Read-Only."
            bByte = 0
         Case 40050
            smsg = "Get New Enum: Couldn't Get Interface."
            bByte = 0
         Case 40051
            smsg = "Assignment To Count Property Not Allowed."
            bByte = 0
         Case 40052
            smsg = "You Must Use Append Chunk To Set Data In A Text Or Image."
            bByte = 0
         Case 40053
            smsg = "Object Collection: Can't Add Non Object Item."
            bByte = 0
         Case 40054
            smsg = "An Invalid Parameter Was Passed."
            bByte = 0
         Case 40055
            smsg = "Invalid Operation."
            bByte = 0
         Case 40056
            smsg = "The Row Has Been Deleted."
            bByte = 0
         Case 40057
            smsg = "An Attempt Was Made To Issue A Select Statement Using Execute."
            bByte = 0
         Case 40058
            smsg = "Can't Update Column, The Result Set Is Read Only."
            bByte = 0
         Case 40059
            smsg = "Cancel Has Been Selected In An ODBC Dialog Requesting Parameters."
            iWarningType = vbInformation
            bByte = 0
         Case 40060
            smsg = "Needs Chunk Required Flags."
            bByte = 0
         Case 40061
            smsg = "Could Not Load Resource Library."
            bByte = 1
         Case 40069
            smsg = "General Client Cursor Error."
            bByte = 0
         Case 40071
            smsg = "The RDO Connection Object Is Not Connected To A Data Source."
            bByte = 1
         Case 40072
            smsg = "The RDO Connection Object Is Already Connected To The Data Source."
            bByte = 0
         Case 40073
            smsg = "The RDO Connection Object Is Busy Connecting " & vbCr _
                   & "To The Data Source. Retry The Selection."
            bByte = 0
         Case 40074
            smsg = "The RDO Query Or RDO Results Set Has No Active Connection Source."
            bByte = 1
         Case 40075
            smsg = "Incorrect Cursor Driver."
            bByte = 0
         Case 40076
            smsg = "This Property Is Currently Read Only."
            iWarningType = vbInformation
            bByte = 0
         Case 40077
            smsg = "The Object Is Already In The Collection."
            iWarningType = vbInformation
            bByte = 0
         Case 40078
            smsg = "Failed To Load RDOCURS.DLL"
            bByte = 1
         Case 40079
            smsg = "Can't Find The Requested Table To Update."
            bByte = 0
         Case 40080, 40081, 40082, 40083, 40085
            smsg = "Invalid RDO/SQL Server Option."
            bByte = 0
         Case 40088
            smsg = "No Open Cursor For Transaction Commit."
            bByte = 0
         Case 40500, 40501, 40502, 40503
            smsg = "Unexpected Internal RDO Error "
            bByte = 1
         Case 40504
            smsg = "Could Not Refresh Controls."
            bByte = 0
         Case 40505
            smsg = "Invalid Property Value."
            bByte = 0
         Case 40506
            smsg = "Invalid Collection Object."
            bByte = 0
         Case 40507
            smsg = "Method Cannot Be Called In RDO's Current State."
            bByte = 0
         Case 40508
            smsg = "One Or More Of The Arguments Is Invalid."
            bByte = 0
         Case 40509
            smsg = "Result Set Is Empty."
            iWarningType = vbInformation
            bByte = 0
         Case 40510
            smsg = "Out Of Memory. Close " & sSysCaption & "."
            bByte = 1
         Case 40511
            smsg = "Result Set Not Available."
            bByte = 0
         Case 40512
            smsg = "The Connection Is Not Open."
            bByte = 1
         Case 40513, 40514
            smsg = "Property Cannot Be Set In RDC's Current State."
            bByte = 0
         Case 40515
            smsg = "Type Mismatch."
            bByte = 0
         Case 40516
            smsg = "Cannot Connect To Remote Data Object."
            bByte = 1
         Case Else
            smsg = "Undocumented Error           "
            bByte = 1
      End Select
      
      Select Case CurrError.Number
         Case 20500 To 20999
            sMsg2 = "Crystal Reports."
         Case 40000 To 40516
            sMsg2 = "SQL Server Warning."
         Case Else
            sMsg2 = sSysCaption & " System."
      End Select
      If iWarningType = vbInformation Then
         smsg = "Notification" & Str(CurrError.Number) & vbCr & smsg & vbCr & sMsg2
      Else
         If bByte = 1 Then
            smsg = "Error" & Str(CurrError.Number) & vbCr & smsg & vbCr & sMsg2
         Else
            smsg = "Warning" & Str(CurrError.Number) & vbCr & smsg & vbCr & sMsg2
         End If
      End If
      MouseCursor 0
      'Show the user and do as required
      MdiSect.Enabled = True
      If bByte = 1 Then
         sErrSev = "16"
         Print #iFreeFile, sDate; sSection; sForm; sUserName; _
            sErrNum; sErrSev; " "; sProcName; " "; Left(CurrError.Description, 20)
         Close iFreeFile
         smsg = smsg & vbCr & "Contact System Administrator"
         MsgBox smsg, vbCritical, frm.Caption
         CloseFiles
      End
   Else
      If iWarningType = vbInformation Then sErrSev = "64" Else sErrSev = "48"
      Print #iFreeFile, sDate; sSection; sForm; sUserName; _
         sErrNum; sErrSev; " "; sProcName; " "; Left(CurrError.Description, 20)
      Close iFreeFile
      sProcName = ""
      MsgBox smsg, iWarningType, frm.Caption
      sErrSev = "48"
      CurrError.Number = 0
   End If
   
End Sub

Public Sub KeyLock(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   'If Combo <> True Then Combo = False
   bUserAction = True
   If KeyAscii = 13 Then SendKeys "{TAB}"
   If KeyAscii > 9 Then KeyAscii = 0
   
End Sub

'Finds and updates Average Cost for a Part

Public Sub AverageCost(sPassedPart As String)
   Dim ActRs As rdoResultset
   Dim cAverageCost As Currency
   
   On Error GoTo modErr1
   sPassedPart = Compress(sPassedPart)
   sSql = "SELECT SUM(INAMT*Abs(INAQTY))/SUM(Abs(INAQTY)) " _
          & "From InvaTable WHERE INAQTY<>0 AND " _
          & "(INPART='" & sPassedPart & "') "
   
   Set ActRs = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not IsNull(ActRs.rdoColumns(0)) Then
      If Val(ActRs.rdoColumns(0)) > 0 Then cAverageCost = ActRs.rdoColumns(0)
   End If
   ActRs.Close
   
   sSql = "UPDATE PartTable SET " _
          & "PAAVGCOST=" & Format(cAverageCost, "#####.0000") _
          & " WHERE PARTREF='" & sPassedPart & "' "
   RdoCon.Execute sSql, rdExecDirect
   
   Set ActRs = Nothing
   Exit Sub
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error Resume Next
   Set ActRs = Nothing
   
End Sub


Public Sub FillCustomers(frm As Form)
   MouseCursor 13
   Dim CstRes As rdoResultset
   On Error GoTo modErr1
   sSql = "Qry_FillCustomers"
   bSqlRows = GetDataSet(CstRes, ES_FORWARD)
   If bSqlRows Then
      With CstRes
         On Error Resume Next
         frm.cmbCst = "" & Trim(!CUNICKNAME)
         frm.txtNme = "" & Trim(!CUNAME)
         frm.lblNum = "" & Trim(!CUNUMBER)
         Do Until .EOF
            '                   frm.cmbCst.AddItem "" & Trim(!CUNICKNAME)
            AddComboStr frm.cmbCst.hWnd, "" & Trim(!CUNICKNAME)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set CstRes = Nothing
   MouseCursor 0
   Exit Sub
   
modErr1:
   sProcName = "fillcustomers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set CstRes = Nothing
   DoModuleErrors frm
   
End Sub

'Use installed query to find all part types 1 thru 3

Public Sub FillPartsBelow4(frm As Form)
   Dim RdoFp4 As rdoResultset
   On Error GoTo modErr1
   sSql = "Qry_PartTypesBelow4"
   bSqlRows = GetDataSet(RdoFp4, ES_FORWARD)
   If bSqlRows Then
      With RdoFp4
         frm.cmbPls = "" & Trim(!PARTNUM)
         Do Until .EOF
            AddComboStr frm.cmbPls.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoFp4 = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillpartsbelow4"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub FindCustomer(frm As Form, sCustomerNickname, Optional bNeedsMore As Byte)
   Dim CusRes As rdoResultset
   On Error GoTo modErr1
   sCustomerNickname = Compress(sCustomerNickname)
   sSql = "Qry_FindCustomer '" & sCustomerNickname & "' "
   bSqlRows = GetDataSet(CusRes)
   If bSqlRows Then
      With CusRes
         On Error Resume Next
         frm.lblCst = "" & Trim(!CUNICKNAME)
         frm.cmbCst = "" & Trim(!CUNICKNAME)
         frm.lblNme = "" & Trim(!CUNAME)
         frm.txtNme = "" & Trim(!CUNAME)
         If bNeedsMore Then
            frm.txtDis = Format(!CUDISCOUNT, "#0.00")
            frm.txtFra = Format(!CUFRTALLOW, "#0.000")
            frm.txtFrd = Format(!CUFRTDAYS, "##0")
         End If
         .Cancel
      End With
   Else
      On Error Resume Next
      frm.lblNme = ""
      frm.txtNme = "*** Customer Wasn't Found ***"
   End If
   Set CusRes = Nothing
   Exit Sub
   
modErr1:
   sProcName = "findcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub SetCrystalAction(frm As Form)
   'fires crystal and sets zoom level
   'if user has selected one
   Dim b As Byte
   Dim FormDriver As String
   Dim FormPort As String
   Dim FormPrinter As String
   
   On Error GoTo modErr1
   MdiSect.crw.ReportTitle = frm.Caption
   MdiSect.crw.WindowTitle = frm.Caption
   If frm.optPrn.Value = True Then
      On Error Resume Next
      Err = 0
      FormPrinter = Trim(frm.lblPrinter)
      If Err > 0 Then FormPrinter = ""
      If FormPrinter = "Default Printer" Then FormPrinter = ""
      If Not bBold Then
         MdiSect.crw.SectionFont(0) = "ALL;;;;N"
      Else
         MdiSect.crw.SectionFont(0) = "ALL;;;;Y"
      End If
      If Len(Trim(FormPrinter)) > 0 Then
         b = GetPrinterPort(FormPrinter, FormDriver, FormPort)
      Else
         FormPrinter = ""
         FormDriver = ""
         FormPort = ""
      End If
      MdiSect.crw.PrinterName = FormPrinter
      MdiSect.crw.PrinterDriver = FormDriver
      MdiSect.crw.PrinterPort = FormPort
      MdiSect.crw.Destination = crptToPrinter
   Else
      MdiSect.crw.Destination = crptToWindow
   End If
   On Error Resume Next
   Err.Clear
   MdiSect.crw.Action = 1
   If Err = 20599 Then
      'Crystal didn't like the DSN, try again
      'using the default (make one if necessary).
      'If it still fails then something else is amiss.
      ' On Error GoTo ModErr1
      sdsn = RegisterSqlDsn("ESI2000")
      MdiSect.crw.Connect = "DSN=" & sdsn & ";UID=" & sSaAdmin & ";PWD=" _
                            & sSaPassword & ";DSQ=" & sDataBase & " "
      SaveSetting "Esi2000", "System", "SqlDsn", sdsn
      MdiSect.crw.Action = 1
   Else
      If Err > 0 Then
         'any other errors
         CurrError.Number = Err.Number
         CurrError.Description = Err.Description
         GoTo Moderr2
      End If
   End If
   If frm.optPrn.Value = False Then
      'Allow for the bug in Crystal that shows the first
      'form full screen
      If bNoCrystal Then
         SendKeys "% R", True
         bNoCrystal = False
      End If
      If iZoomLevel > 0 Then MdiSect.crw.PageZoom (iZoomLevel)
   End If
   Exit Sub
   
modErr1:
   sProcName = "setcrystalact"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
Moderr2:
   DoModuleErrors frm
   
End Sub

Public Sub GetCrystalZoom()
   Dim sCrystalBold As String
   Dim sCrystalZoom As String
   sCrystalZoom = GetSetting("Esi2000", "System", "ReportZoom", sCrystalZoom)
   iZoomLevel = Val(sCrystalZoom)
   
   sCrystalBold = GetSetting("Esi2000", "System", "ReportBold", sCrystalBold)
   bBold = Val(sCrystalBold)
   
End Sub

'Test to see if we are in VB or user mode

Public Function RunningInIDE() As Boolean
    'Check to see where the program is running.
    'Assume that we are not running in VB5 for now
    'Calling Debug is ignored except in VB, so will produce
    'an error...
    On Error GoTo ModErr
    Debug.Print(1 / 0)
    'nope, not Vb
    RunningInIDE = False
    Exit Function

ModErr:
    On Error GoTo 0
    'yep, it's VB
    RunningInIDE = True

End Function


Public Sub Notes()
   
   '  Connect Property and OpenConnection Example: DSN-Less Connection Using OpenConnection
   '  The following example establishes a DSN-less ODBC connection using the OpenConnection method against the default rdoEnvironment. In this case the example prints the resulting Connect property to the Immediate window.
   
   '  Dim en As rdoEnvironment
   '  Dim cn As rdoConnection
   
   '  Set en = rdoEnvironments(0)
   '  Set cn = en.OpenConnection(dsName:="", _
   '      Prompt:=rdDriverNoPrompt, _
   '      Connect:="uid=;pwd=;driver={SQL Server};" _
   '          & "server=SEQUEL;database=pubs;")
   '
   '6/16/98    Added EsiKeyBd Class Module
   
End Sub

'6/9/99-Removed references to ACCRA
'Use the default ESI2000 DSN if one hasn't been registered
'in ES/2000

Public Sub GetCrystalConnect()
   Dim sConn As String
   MdiSect.crw.Connect = "DSN=" & sdsn & ";UID=" & sSaAdmin & ";PWD=" _
                         & sSaPassword & ";DSQ=" & sDataBase
   MdiSect.crw.WindowBorderStyle = crptSizable
   MdiSect.crw.WindowControlBox = True
   MdiSect.crw.WindowMaxButton = True
   MdiSect.crw.WindowMinButton = True
   MdiSect.crw.WindowShowCancelBtn = True
   MdiSect.crw.WindowShowCloseBtn = True
   MdiSect.crw.WindowShowExportBtn = True
   MdiSect.crw.WindowShowGroupTree = False
   MdiSect.crw.WindowShowNavigationCtls = True
   MdiSect.crw.WindowShowPrintBtn = True
   MdiSect.crw.WindowShowPrintSetupBtn = True
   MdiSect.crw.WindowShowRefreshBtn = False
   MdiSect.crw.WindowShowZoomCtl = True
   MdiSect.crw.WindowShowSearchBtn = False
   MdiSect.crw.WindowAllowDrillDown = False
   '  MdiSect.Crw.ProgressDialog = False
   Exit Sub
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error GoTo 0
   
End Sub


Public Sub GetCurrentSelections()
   Cur.CurrentPart = GetSetting("Esi2000", "Current", "Part", Cur.CurrentPart)
   Cur.CurrentVendor = GetSetting("Esi2000", "Current", "Vendor", Cur.CurrentVendor)
   Cur.CurrentCustomer = GetSetting("Esi2000", "Current", "Customer", Cur.CurrentCustomer)
   Cur.CurrentRegion = GetSetting("Esi2000", "Current", "Region", Cur.CurrentRegion)
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", bEnterAsTab)
   
End Sub

Public Sub SaveCurrentSelections()
   SaveSetting "Esi2000", "Current", "Part", Cur.CurrentPart
   SaveSetting "Esi2000", "Current", "Vendor", Cur.CurrentVendor
   SaveSetting "Esi2000", "Current", "Customer", Cur.CurrentCustomer
   SaveSetting "Esi2000", "Current", "Region", Cur.CurrentRegion
End Sub

'Sets the form size to allow Resize.ocx to do
'whatever it does...
'See ES_RESIZE, ES_DONTRESIZE Constansts

Public Sub SetFormSize(frm As Form)
   Dim cNewSize As Currency
   
   On Error Resume Next
   If bResize = 1 Then
      If lScreenWidth > 9999 Then '640X480
         If lScreenWidth < 14000 Then '800X600
            cNewSize = 1.05
         Else
            cNewSize = 1.1 '1024X768
         End If
         frm.Height = frm.Height * cNewSize
         frm.Width = frm.Width * cNewSize
      End If
   Else
      frm.ReSize1.Enabled = False
   End If
   
End Sub



Public Sub SelectFormat(frm As Form)
   'Selects all of the text in fixed length text boxes
   'and combo boxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   'Allow for the possibility that the control has been
   'disabled
   
   On Error Resume Next
   frm.ActiveControl.SelStart = 0
   frm.ActiveControl.SelLength = Len(frm.ActiveControl)
   
End Sub




Public Sub FillDivisions(frm As Form)
   Dim RdoDiv As rdoResultset
   On Error GoTo modErr1
   RdoCon.QueryTimeout = 40
   Set RdoDiv = RdoCon.OpenResultset("Qry_FillDivisions", rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoDiv.BOF And Not RdoDiv.EOF Then
      With RdoDiv
         Do Until .EOF
            'frm.cmbDiv.AddItem "" & Trim(!DIVREF)
            AddComboStr frm.cmbDiv.hWnd, "" & Trim(!DIVREF)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   On Error Resume Next
   Set RdoDiv = Nothing
   Exit Sub
   
modErr1:
   sProcName = "filldivisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub FillRegions(frm As Form)
   Dim RdoReg As rdoResultset
   On Error GoTo modErr1
   RdoCon.QueryTimeout = 40
   Set RdoReg = RdoCon.OpenResultset("Qry_FillRegions", rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoReg.BOF Then
      With RdoReg
         Do Until .EOF
            'frm.cmbReg.AddItem "" & Trim(!REGREF)
            AddComboStr frm.cmbReg.hWnd, "" & Trim(!REGREF)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoReg = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillregions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub FillProductCodes(frm As Form)
   Dim RdoCde As rdoResultset
   On Error GoTo modErr1
   RdoCon.QueryTimeout = 40
   Set RdoCde = RdoCon.OpenResultset("Qry_FillProductCodes", rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoCde.BOF Then
      With RdoCde
         Do Until .EOF
            AddComboStr frm.cmbCde.hWnd, "" & Trim(!PCCODE)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCde = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillproductcodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub FillProductClasses(frm As Form)
   Dim RdoCls As rdoResultset
   On Error GoTo modErr1
   RdoCon.QueryTimeout = 40
   Set RdoCls = RdoCon.OpenResultset("Qry_FillProductClasses", rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoCls.BOF Then
      With RdoCls
         Do Until .EOF
            AddComboStr frm.cmbCls.hWnd, "" & Trim(!CCCODE)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoCls = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillproductclasses"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

Public Sub FillTerms(frm As Form)
   Dim RdoTrm As rdoResultset
   On Error GoTo modErr1
   RdoCon.QueryTimeout = 40
   Set RdoTrm = RdoCon.OpenResultset("SELECT TRMREF FROM StrmTable", rdOpenForwardOnly, rdConcurReadOnly)
   If Not RdoTrm.BOF Then
      With RdoTrm
         Do Until .EOF
            'frm.cmbTrm.AddItem "" & Trim(!TRMREF)
            AddComboStr frm.cmbTrm.hWnd, "" & Trim(!TRMREF)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoTrm = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

'Load and show the calendar from the combo dropdown

Public Sub ShowCalendar(frm As Form, Optional iAdjust As Integer)
   Dim i As Integer
   Dim iAdder As Integer
   Dim lLeft As Long
   Dim sDate As String
   
   'Blast thru any errors
   On Error Resume Next
   'See if there is a date in the combo
   If Len(Trim(frm.ActiveControl.Text)) Then
      sDate = frm.ActiveControl.Text
   Else
      sDate = Format(Now, "mm/dd/yy")
   End If
   
   'Larry doen't like it-Let's do this instead (2/24/99)
   
   'If it's loaded and the user clicks it, then hide it
   '    If bCalendar Then
   '        If frm.ActiveControl.ListCount > 0 Then
   '             For i = frm.ActiveControl.ListCount - 1 To 1 Step -1
   '                 frm.ActiveControl.RemoveItem frm.ActiveControl.List(i)
   '             Next
   '        Else
   '            frm.ActiveControl.AddItem frm.ActiveControl.Text
   '        End If
   '        frm.ActiveControl = sDate
   '        bCalendar = False
   '        Calendar.Hide
   '    Else
   'Or find out where to put it and show it
   If iBarOnTop = 0 Then
      lLeft = frm.Left + frm.ActiveControl.Left
   Else
      'lLeft = 0
      lLeft = frm.ActiveControl.Left
   End If
   
   If iBarOnTop = 0 Then If lLeft > 6000 Then lLeft = lLeft - 1095
   If (lLeft + Calendar.Width) > (MdiSect.Width - 600) Then lLeft = lLeft - (Calendar.Width - frm.ActiveControl.Width + 300)
   
   If iBarOnTop = 0 Then
      Calendar.Move MdiSect.SideBar.Width + lLeft, frm.Top + (frm.ActiveControl.Top + frm.ActiveControl.Height + 1000 + iAdjust)
   Else
      Calendar.Move lLeft, frm.Top + (frm.ActiveControl.Top + MdiSect.TopBar.Height + frm.ActiveControl.Height + 1000 + iAdjust)
   End If
   
   bCalendar = True
   If Len(sDate) Then Calendar.Calendar1.Value = CDate(sDate)
   Calendar.Show
   'refresh it so that it doesn't blink out
   Calendar.Calendar1.Refresh
   '  End If
   
End Sub


'Converts real hours like 8.3 to time like 08:18
'Note systax - Pass a number, returns a string or variant

Public Function ConvertHours(cTime As Currency) As String
   Dim cChg As Currency
   Dim sNewTime As String * 2
   Dim sTime As String
   
   On Error GoTo modErr1
   sTime = Format$(cTime * 100, "00:00")
   cChg = Val(Right(sTime, 2)) * 6
   sNewTime = Format$(cChg, "000")
   ConvertHours = Left(sTime, 3) & sNewTime
   Exit Function
   
modErr1:
   Resume Moderr2
Moderr2:
   ConvertHours = "00:00"
   On Error GoTo 0
   
End Function

'gets password incription (self documenting)

Public Function GetPassword(sPassword As String) As String
   Dim k As Integer
   Dim i As Integer
   Dim n As Integer
   Dim sNewPw As String
   
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

Public Function SetPassword(sPassword As String) As String
   Dim k As Integer
   Dim i As Integer
   Dim n As Integer
   Dim sNewPw As String
   
   On Error Resume Next
   k = Len(Trim$(sPassword))
   For i = k To 1 Step -1
      sNewPw = sNewPw & Mid$(sPassword, i, 1)
   Next
   n = 79
   For i = 1 To k
      n = n + 1
      Mid$(sNewPw, i, 1) = Chr$(Asc(Mid$(sNewPw, i, 1)) + n)
   Next
   SetPassword = sNewPw
   
End Function

'Retrieves the list of most recent selections from
'the registry and files the ActiveBar.
'Might as well do some other load stuff too

Public Sub GetRecentList(sMdiSect As String)
   Dim i As Integer
   On Error Resume Next
   MdiSect.TmePanel = Format(Time, "h:mm AM/PM")
   'Initialize the rdoEngine
   'Set RdoEnv = rdoEnvironments(0)
   'User
   If Cur.CurrentUser = "" Then Cur.CurrentUser = GetSetting("Esi2000", "system", "UserId", Cur.CurrentUser)
   'Server
   sServer = UCase$(GetSetting("Esi2000", "System", "ServerId", sServer))
   'DSN for Crystal - Make one if required
   sdsn = GetSetting("Esi2000", "System", "SqlDsn", sdsn)
   If Trim(sdsn) = "" Then
      sdsn = RegisterSqlDsn("ESI2000")
      If Trim(sdsn) <> "" Then SaveSetting "Esi2000", "System", "SqlDsn", sdsn
   End If
   
   On Error GoTo modErr1
   bNoCrystal = True
   For i = 0 To 4
      sRecent(i%) = GetSetting("Esi2000", sMdiSect, "Recent" & Trim(Str(i)), sRecent(i))
      If Len(Trim(sRecent(i))) < 2 Then
         'Nothing there and hide it
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(Str(i))).Visible = False
      Else
         'There is an entry and show it
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(Str(i))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(Str(i))).Caption = sRecent(i)
      End If
   Next
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

'Set common MdiForm activation methods
'Give SQL Server a chance to check us in

Public Sub ActivateSection(sCurrentSection As String)
   'initialize the section
   Dim bOpenForm As Byte
   On Error Resume Next
   MdiSect.BotPanel = "Initializing."
   OpenSqlServer False
   'Make sure that SQL Server is going to connect or fail
   If RdoCon.StillConnecting Then
      MdiSect.Enabled = False
      Sleep 1000
   End If
   MdiSect.Enabled = True
   bUserAction = True
   MdiSect.BotPanel = "Ready.."
   Sleep 500
   MdiSect.BotPanel.FontItalic = False
   MdiSect.BotPanel = MdiSect.Caption
   MdiSect.SetFocus
   bOpenForm = GetSetting("Esi2000", "System", "Reopenforms", bOpenForm)
   If bOpenForm Then
      sSelected = GetSetting("Esi2000", sCurrentSection, "LastBox", sSelected)
      If Len(Trim(sSelected)) Then OpenFavorite sSelected
   End If
   If sDataBase <> "Esi2000Db" Then
      bSecSet = 0
      User.Group1 = 1
      User.Group2 = 1
      User.Group3 = 1
      User.Group4 = 1
      User.Group5 = 1
      User.Group6 = 1
   End If
   MouseCursor 0
   
End Sub

'Standard Unload as save for all sections
'MdiSect QueryUnload Event

Public Sub UnLoadSection(sMdiSect As String, sThisSection As String)
   Dim i As Integer
   'tell mom we are not here
   SaveSetting "Esi2000", "Sections", sMdiSect, 0
   'save Recent list
   SaveSetting "Esi2000", sThisSection, "LastBox", sSession(0)
   For i = 0 To 4
      SaveSetting "Esi2000", sThisSection, "Recent" & Trim(Str(i)), Trim(MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(Str(i))).Caption)
   Next
   
End Sub

'Standardize then Resize of the MdiForm

Public Sub ResizeSection()
   iBarOnTop = GetSetting("Esi2000", "Programs", "BarOnTop", iBarOnTop)
   If iBarOnTop = 1 Then
      MdiSect.SideBar.Visible = False
      MdiSect.TopBar.Visible = True
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
   Else
      MdiSect.SideBar.Visible = True
      MdiSect.TopBar.Visible = False
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
   End If
   MdiSect.TmePanel.Left = (MdiSect.BotPanel.Width - 850)
   MdiSect.OvrPanel.Left = (MdiSect.BotPanel.Width - 1650)
   MdiSect.lblShop.Left = (MdiSect.BotPanel.Width - 3400)
   
End Sub

'Standardize the Sub Main procedure

Public Sub MainLoad(sEsiSection As String)
   Dim i As Integer
   MouseCursor 11
   
   'Trap App to see if MOM is watching. Code for CurDir
   'to allow testing and programming
    bUserAction = RunningInIDE()
   If Not bUserAction Then
      i = GetSetting("Esi2000", "sections", "EsiOpen", i)
      If i = 0 Then
         Y = False
         MouseCursor 0
         Awarn.Show 1
         Do Until Y
         Loop
      End
   End If
Else
   Cur.CurrentUser = "ADMINISTRATOR"
   'Cur.CurrentUser = "COLINS"
End If
sFilePath = GetSetting("Esi2000", "System", "FilePath", sFilePath)
'tell mom we are here
If i = 1 Then SaveSetting "Esi2000", "Sections", sEsiSection, 1
bUserAction = True
MdiSect.BotPanel.FontItalic = True

End Sub


'Pickup the sa and sa password if recorded
'Refined 3/5/02

Public Function GetSysLogon(sSaPassword As Byte) As String
   Dim a As Integer
   Dim i As Integer
   Dim sTest As String
   Dim sNewString As String
   Dim sPassword As String
   
   If sSaPassword Then
      GetSysLogon = GetSetting("UserObjects", "System", "NoReg", GetSysLogon)
      If Trim(GetSysLogon) = "" Then GetSysLogon = "sa"
   Else
      sPassword = GetSetting("SysCan", "System", "RegOne", sPassword)
      sPassword = Trim(sPassword)
      If sPassword <> "" Then
         i = Len(sPassword)
         If i > 5 Then
            sPassword = Mid(sPassword, 4, i - 5)
         End If
      End If
      GetSysLogon = sPassword
   End If
   
End Function

'Fills state codes

Public Sub FillStates(frm As Form)
   Dim RdoSte As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT STATECODE,STATEDEFAULT FROM CsteTable"
   bSqlRows = GetDataSet(RdoSte)
   If bSqlRows Then
      With RdoSte
         On Error Resume Next
         Do Until .EOF
            'for vendors..
            If frm.Name = "diaPvndr" Then
               'frm.cmbPste.AddItem "" & Trim(!STATECODE)
               AddComboStr frm.cmbPste.hWnd, "" & Trim(!STATECODE)
               If !STATEDEFAULT = 1 Then
                  frm.cmbSte = "" & Trim(!STATECODE)
               End If
            End If
            If frm.Name = "diaCcust" Then
               'frm.cmbStSte.AddItem "" & Trim(!STATECODE)
               'frm.cmbBtSte.AddItem "" & Trim(!STATECODE)
               AddComboStr frm.cmbStSte.hWnd, "" & Trim(!STATECODE)
               AddComboStr frm.cmbBtSte.hWnd, "" & Trim(!STATECODE)
               If !STATEDEFAULT = 1 Then
                  frm.cmbStSte = "" & Trim(!STATECODE)
                  frm.cmbBtSte = "" & Trim(!STATECODE)
               End If
            End If
            'frm.cmbSte.AddItem "" & Trim(!STATECODE)
            AddComboStr frm.cmbSte.hWnd, "" & Trim(!STATECODE)
            If !STATEDEFAULT = 1 Then
               frm.cmbSte = "" & Trim(!STATECODE)
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoSte = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillstates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Moderr2
Moderr2:
   If Left(CurrError.Description, 5) = "S0002" Then
      CurrError.Number = 0
      CurrError.Description = ""
   Else
      DoModuleErrors frm
   End If
   
End Sub


'Remove spaces (32), dashes (45) and tabs (9) from indexed fields to
'avoid duplicate entries
'Optionally trim the length of the entry
'8/14/99 added optional ES_IGNOREDASHES to compress leaving dashes
'note requires length if used

Public Function Compress(TestNo As Variant, Optional iLength As Integer, Optional bIgnoreDashes As Byte) As String
   Dim a As Integer
   Dim k As Integer
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(TestNo)
   a = Len(PartNo)
   If a > 0 Then
      For k = 1 To a
         If bIgnoreDashes Then
            If Mid$(PartNo, k, 1) <> Chr$(32) And Mid$(PartNo, k, 1) <> Chr$(9) _
                    And Mid$(PartNo, k, 1) <> Chr$(39) Then
               NewPart = NewPart & Mid$(PartNo, k, 1)
            End If
         Else
            If Mid$(PartNo, k, 1) <> Chr$(45) And Mid$(PartNo, k, 1) <> Chr$(32) _
                    And Mid$(PartNo, k, 1) <> Chr$(9) And Mid$(PartNo, k, 1) <> Chr$(39) Then
               NewPart = NewPart & Mid$(PartNo, k, 1)
            End If
         End If
      Next
   End If
   If iLength > 0 Then
      If Len(NewPart) > iLength Then
         Beep
         NewPart = Left$(NewPart, iLength)
      End If
   End If
   Compress = NewPart
   Exit Function
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error Resume Next
   Compress = TestNo
   
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


'Replaces in all but EsiFina 1/22/01
'New routine is AutoFormatControls

Public Sub FormatFormControls(frm As Form)
   ' //Need the following in case of a untrapped Control Array.
   ' Manual Code those from Module Procedures.
   'Dim bByte As Byte
   'Dim i As Integer
   'Dim A As Integer
   'Dim b As Integer
   'Dim C As Integer
   'Dim n As Integer
   '    A = -1
   '    b = -1
   '
   '    Erase ESI_txtGotFocus
   '    Erase ESI_txtKeyPress
   '    Erase ESI_txtKeyDown
   '
   '     'Have to allow for arrays, etc-blast thru
   '      On Error Resume Next
   '        For i = 0 To frm.Controls.Count - 1
   '           'Part of an Array or label (z1(n))?
   '          C = frm.Controls(i).Index
   '          If TypeOf frm.Controls(i) Is SSRibbon Then
   '              If frm.Controls(i).Name = "ShowPrinters" Then
   '                  Set ESI_cmdShowPrint.esCmdClick = frm.Controls(i)
   '              End If
   '          End If
   '          If Err > 0 And (TypeOf frm.Controls(i) Is TextBox Or _
   '              TypeOf frm.Controls(i) Is ComboBox Or TypeOf frm.Controls(i) Is MaskEdBox) Then
   '                Err = 0
   '                A = A + 1
   '                ReDim Preserve ESI_txtKeyPress(A) As New EsiKeyBd
   '                If frm.Controls(i).Tag <> "9" Then
   '                    b = b + 1
   '                    ReDim Preserve ESI_txtGotFocus(b) As New EsiKeyBd
   '                    ReDim Preserve ESI_txtKeyDown(b) As New EsiKeyBd
   '                End If
   '                If TypeOf frm.Controls(i) Is MaskEdBox Then
   '                    Set ESI_txtGotFocus(b).esMskGotFocus = frm.Controls(i)
   '                    Set ESI_txtKeyDown(b).esMskKeyDown = frm.Controls(i)
   '                    Set ESI_txtKeyPress(A).esMskKeyValue = frm.Controls(i)
   '                End If
   '                If TypeOf frm.Controls(i) Is TextBox Then
   '                  bByte = True
   '                  Select Case Val(frm.Controls(i).Tag)
   '                      Case 1
   '                          Set ESI_txtKeyPress(A).esTxtKeyValue = frm.Controls(i)
   '                      Case 3
   '                          Set ESI_txtKeyPress(A).esTxtKeyCase = frm.Controls(i)
   '                      Case 4
   '                          Set ESI_txtKeyPress(A).esTxtKeyDate = frm.Controls(i)
   '                      Case 5
   '                          Set ESI_txtKeyPress(A).esTxtKeyTime = frm.Controls(i)
   '                      Case 9
   '                          Set ESI_txtKeyPress(A).esTxtKeyMemo = frm.Controls(i)
   '                          bByte = False
   '                      Case Else
   '                          Set ESI_txtKeyPress(A).esTxtKeyCheck = frm.Controls(i)
   '                  End Select
   '                  If bByte Then
   '                      Set ESI_txtGotFocus(b).esTxtGotFocus = frm.Controls(i)
   '                      Set ESI_txtKeyDown(b).estxtKeyDown = frm.Controls(i)
   '                  End If
   '               Else
   '                  If TypeOf frm.Controls(i) Is ComboBox Then
   '                    Set ESI_txtGotFocus(b).esCmbGotfocus = frm.Controls(i)
   '                    Select Case Val(frm.Controls(i).Tag)
   '                      Case 1
   '                          Set ESI_txtKeyPress(A).esCmbKeyValue = frm.Controls(i)
   '                      Case 4
   '                          Set ESI_txtKeyPress(A).esCmbKeyDate = frm.Controls(i)
   '                      Case 8
   '                          Set ESI_txtKeyPress(A).esCmbKeylock = frm.Controls(i)
   '                          frm.Controls(i).ForeColor = ES_BLUE
   '                      Case Else
   '                          Set ESI_txtKeyPress(A).esCmbKeyCase = frm.Controls(i)
   '                    End Select
   '                  End If
   '              End If
   '            End If
   '        Next
   
End Sub

'Removes alot of code for this and makes it consistant

Public Sub CancelTrans()
   'Maybe its not an mdi form
   On Error GoTo modErr1
   MsgBox "Transaction Canceled By The User.", _
      vbInformation, MdiSect.ActiveForm.Caption
   Exit Sub
   
modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   MsgBox "Transaction Canceled By The User.", _
      vbInformation, sSysCaption
   
End Sub

'Standard procedure for receiving resultsets
'bSqlRows = GetDataSet (RdoRes, ES_FORWARD)
'See also GetQuerySet for VB built queries

Public Function GetDataSet(RdoDataSet As rdoResultset, Optional iCursorType As Integer) As Byte
   ' Use local error Trapping
   RdoCon.QueryTimeout = 40
   If iCursorType = ES_FORWARD Then
      'Forward only "cursor" (not a cursor)
      Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
      If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
         GetDataSet = 1
      Else
         GetDataSet = 0
      End If
   Else
      If iCursorType = ES_KEYSET Then
         'Keyset cursor for Editing
         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenKeyset, rdConcurRowVer)
         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
            GetDataSet = 1
         Else
            GetDataSet = False
         End If
      ElseIf iCursorType = ES_DYNAMIC Then
         'Dynamic
         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenDynamic, rdConcurRowVer)
         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
            GetDataSet = 1
         Else
            GetDataSet = 0
         End If
      ElseIf iCursorType = ES_STATIC Then
         'Static Cursor - Note: Needed for BLOBS
         Set RdoDataSet = RdoCon.OpenResultset(sSql, rdOpenStatic, rdConcurReadOnly)
         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
            GetDataSet = 1
         Else
            GetDataSet = 0
         End If
      End If
      If Err > 0 Then GetDataSet = 0
   End If
   
End Function

'Standard procedure for receiving resultsets
'bSqlRows = GetQuerySet (RdoRes, RdoQry, ES_FORWARD)
'See also GetDataSet for general queries

Public Function GetQuerySet(RdoDataSet As rdoResultset, RdoQueryDef As rdoQuery, Optional iCursorType As Integer) As Byte
   ' Use local error Trapping
   RdoCon.QueryTimeout = 40
   If iCursorType = ES_FORWARD Then
      'Forward only "cursor" (not a real cursor)
      Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
      If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
         GetQuerySet = 1
      Else
         GetQuerySet = 0
      End If
   Else
      'Keyset cursor for Editing
      If iCursorType = ES_KEYSET Then
         Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenKeyset, rdConcurRowVer)
         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
            GetQuerySet = 1
         Else
            GetQuerySet = 0
         End If
      Else
         'Static Cursor - Note: Needed for BLOBS
         Set RdoDataSet = RdoQueryDef.OpenResultset(rdOpenStatic, rdConcurReadOnly)
         If Not RdoDataSet.BOF And Not RdoDataSet.EOF Then
            GetQuerySet = 1
         Else
            GetQuerySet = 0
         End If
      End If
      If Err > 0 Then GetQuerySet = 0
   End If
   
End Function

Public Sub FormInitialize()
   Dim a As Long
   Dim iToolsCount As Integer
   On Error GoTo modErr1
   bUserAction = True
   'Bring some order to the user. Like it or not 10/2/01
   MdiSect.ActiveBar1.BackColor = Es_FormBackColor
   For iToolsCount = 0 To MdiSect.Controls.Count - 1
      If TypeOf MdiSect.Controls(iToolsCount) Is SSRibbon Then
         MdiSect.Controls(iToolsCount).BackColor = Es_FormBackColor
         MdiSect.Controls(iToolsCount).MousePointer = 99
      Else
         If TypeOf MdiSect.Controls(iToolsCount) Is Label Then
            MdiSect.Controls(iToolsCount).BackColor = Es_FormBackColor
            MdiSect.Controls(iToolsCount).ForeColor = Es_TextForeColor
         Else
            If TypeOf MdiSect.Controls(iToolsCount) Is SSPanel Then
               MdiSect.Controls(iToolsCount).BackColor = Es_FormBackColor
               MdiSect.Controls(iToolsCount).ForeColor = Es_TextForeColor
            End If
         End If
      End If
   Next
   MdiSect.TmePanel.Left = (MdiSect.BotPanel.Width - 850)
   MdiSect.OvrPanel.Left = (MdiSect.BotPanel.Width - 1650)
   MdiSect.lblShop.Left = (MdiSect.BotPanel.Width - 3400)
   MdiSect.Left = 10
   MdiSect.Top = 10
   MdiSect.Width = Screen.Width - 100
   MdiSect.Height = Screen.Height - 100
   'MdiSect.ReSize1.FormMinHeight = Screen.Height / 1.5
   'MdiSect.ReSize1.FormMinWidth = Screen.Width / 1.5
   MdiSect.crw.WindowState = crptNormal
   Exit Sub
   
modErr1:
   On Error GoTo 0
End Sub


Public Sub ActivityDocument()
   'Activity Codes
   'Inventory activity (InvaTable)
   '   1 = Beginning Balance (first established)
   '  19 = Manual Adjustments
   'Pick Types - some not used yet
   '   9 = Pick Request (database default)
   '  10 = Actual Pick
   '  11 = Pick On Dock
   '  12 = Canceled Pick Request
   '  13 = Pick Surplus
   '  20 = Not used
   '  21 = Restocked Pick item
   '  22 = Scrapped Pick item
   '  23 = Pick Substitute
   '  27 = Pick From Freight
   
   'PO Items - some not used yet
   '  14 = Open PO Item
   '  15 = PO Receipt
   '  16 = Canceled PO Item
   '  17 = Invoiced PO Item
   '  18 = ON DOCK
   
   'Shipped Items
   '   3 = Shipment Item
   '   4 = Returned Item
   '   5 = Canceled So Item (Shipped)
   
   'Pack Slips
   '  25 = Packing Slip (inv out)
   '  33 = Packing Slip Canceled (inv in)
   
   'MO's
   '   6 = Completed MO
   '   7 = Closed MO Revised from 6 to 7
   '  38 = Canceled Mo Completion (inventory out)
   
   '   Insert routine
   'Update Part Qoh
   '     sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Val(vItems(i, 2)) & " " _
   '         & "WHERE PARTREF='" & LTrim(Str(vItems(i, 3))) & "' "
   '     RdoCon.Execute sSql, rdExecDirect
   '     AverageCost LTrim(Str(vItems(i, 3)))
   
   'Add to Activity
   '    sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
   '        & "INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT) " _
   '        & "VALUES(25,'" & vItems(i, 3) & "','PACKING SLIP'," _
   '        & "'" & vItems(i, 0) & Trim(vItems(i, 1)) & "'," _
   '        & "'" & Format(Now, "mm/dd/yy") & "'," & Val(vItems(i, 2)) & "," _
   '        & Val(vItems(i, 4)) & ",'" & sCreditAcct & "','" & sDebitAcct & "')"
   '    RdoCon.Execute sSql, rdExecDirect
   '----------
End Sub

'Format for name fields etc
'See constant ES_FIRSTWORD for First word only

Public Function StrCase(sTextStr As Variant, Optional bTextOption As Byte)
   Dim iStrLen As Integer
   Dim sNewStr As String
   If bAutoCaps = 1 Then
      StrCase = sTextStr
      Exit Function
   End If
   
   sNewStr = Trim$(sTextStr)
   iStrLen = Len(sNewStr)
   If iStrLen > 1 Then
      If Asc(Mid(sNewStr, 2, 1)) > 64 And Asc(Mid(sNewStr, 2)) < 90 Then
         StrCase = sNewStr
         Exit Function
      End If
   End If
   If iStrLen > 0 Then
      If bTextOption = ES_FIRSTWORD Then
         'First word of string capitalized
         Mid(sNewStr, 1, 1) = UCase(Left$(sNewStr, 1))
      Else
         'First letter of each word capitalized
         sNewStr = StrConv(sNewStr, vbProperCase)
      End If
   End If
   StrCase = sNewStr
   
End Function


Public Sub GetSystemMessage()
   Static sOldMessage As String
   Dim b As Byte
   Dim RdoMsg As rdoResultset
   
   MdiSect.tmr4.Enabled = False
   On Error GoTo modErr1
   sSql = "Qry_GetSysMessage"
   Set RdoMsg = RdoCon.OpenResultset(sSql, rdOpenForwardOnly)
   If Not RdoMsg.BOF And Not RdoMsg.EOF Then
      With RdoMsg
         MdiSect.SystemMsg.ForeColor = ES_RED
         If sOldMessage <> "" & Trim(!ALERTMSG) Then
            Beep
            MdiSect.SystemMsg = "" & Trim(!ALERTMSG)
            sOldMessage = "" & Trim(MdiSect.SystemMsg)
            If Len(sOldMessage) > 24 Then
               MdiSect.SystemMsg.Alignment = 0
            Else
               MdiSect.SystemMsg.Alignment = 2
            End If
         End If
         .Cancel
      End With
   End If
   Set RdoMsg = Nothing
   MdiSect.tmr4.Enabled = True
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub


'Find the post date for a period
'txtDte = GetPostDate(Me, txtDte)

Public Function GetPostDate(frm As Form, sDate As String) As String
   Dim RdoDte As rdoResultset
   Dim b As Byte
   Dim i As Integer
   
   On Error GoTo modErr1
   For i = 1 To 13
      sSql = "SELECT FYYEAR,FYPERSTART" & Trim(Str(i)) & "," _
             & "FYPEREND" & Trim(Str(i)) & " FROM GlfyTable WHERE ('" _
             & sDate & "' BETWEEN FYPERSTART" & Trim(Str(i)) _
             & " AND FYPEREND" & Trim(Str(i)) & ") "
      bSqlRows = GetDataSet(RdoDte, ES_FORWARD)
      If bSqlRows Then
         GetPostDate = Format(RdoDte.rdoColumns(2), "mm/dd/yy")
         Exit For
      End If
   Next
   If GetPostDate = "" Then
      MsgBox "No Posting Date In The Period Selected.", _
         vbExclamation, frm.Caption
   End If
   Exit Function
   
modErr1:
   'if the table isn't there, forget it
   If Err.Number <> 40002 Then
      sProcName = "getpostdate"
      CurrError.Number = Err.Number
      CurrError.Description = Err.Description
      DoModuleErrors frm
   End If
   
End Function


'retrieve columns form printed forms with User headings
'syntax bReport = GetPrintedForm("Pack Slip")

Function GetPrintedForm(sForm As String) As Byte
   Dim RdoFrm As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT PreRecord,PrePackSlip,PreInvoice," _
          & "PrePurchaseOrder,PreStateMent FROM " _
          & "Preferences WHERE PreRecord=1"
   bSqlRows = GetDataSet(RdoFrm, ES_FORWARD)
   If bSqlRows Then
      With RdoFrm
         Select Case UCase$(Compress(sForm))
            Case "PACKSLIP"
               GetPrintedForm = .rdoColumns(1)
            Case "INVOICE"
               GetPrintedForm = .rdoColumns(2)
            Case "PURCHASEORDER"
               GetPrintedForm = .rdoColumns(3)
            Case "STATEMENT"
               GetPrintedForm = .rdoColumns(4)
            Case Else
               GetPrintedForm = 0
         End Select
      End With
   End If
   Set RdoFrm = Nothing
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error GoTo 0
   
End Function

'Use local errors
'Execute direct for SQL Server...note stop on "'" (ANSI 39)

Public Function CheckComments(sComments As String) As String
   Dim a As Integer
   Dim i As Integer
   a = Len(Trim(sComments))
   If a > 0 Then
      For i = 1 To a
         If Mid(sComments, i, 1) = Chr$(39) Then
            Mid(sComments, i, 1) = Chr$(180)
         End If
      Next
   End If
   CheckComments = sComments
   
End Function

Public Function GetTimeOut(sLastTime As String) As String
   On Error GoTo modErr1
   GetTimeOut = "Last Access " & sLastTime & ", Timeout " _
                & Format(Time, "hh:mm AM/PM") & vbCrLf _
                & "Normal Database Connection Timeout." & vbCrLf _
                & "Reconnect To Service?"
   Exit Function
   
modErr1:
   Resume Next
   
End Function


'For numbers Not used yet. Use Str(lNumber)

Public Sub AddComboNum(lhWnd As Long, lNumber As Long)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, _
      ByVal lNumber
   
End Sub

'Need to watch what the enter

Public Function CheckValidColumn(sColumn As Variant) As Boolean
   Dim a As Integer
   Dim k As Integer
   Dim g As Byte
   Dim b As Byte
   Dim sRefCol As String
   
   On Error GoTo modErr1
   sRefCol = Trim(sColumn)
   a = Len(sRefCol)
   If a > 0 Then
      For k = 1 To a
         If Mid$(sRefCol, k, 1) <> Chr$(32) Then
            If Mid$(sRefCol, k, 1) < Chr$(46) Then
               b = 1
               g = Asc(Mid$(sRefCol, k, 1))
               Exit For
            End If
         End If
      Next
   End If
   If b = 1 Then
      MsgBox "This Entry Includes and Invalid Character " & Chr$(g) & ".", _
         vbExclamation, sSysCaption
      CheckValidColumn = False
   Else
      CheckValidColumn = True
   End If
   Exit Function
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error Resume Next
   CheckValidColumn = False
   
   
End Function


'Allows selection of printers for individual reports
'Crystal requires this stuff

Public Function GetPrinterPort(devPrinter As String, devDriver As String, devPort As String) As Byte
   Dim X As Printer
   For Each X In Printers
      If Trim(X.DeviceName) = devPrinter Then
         devDriver = X.DriverName
         devPort = X.Port
         Exit For
      End If
   Next
   
End Function

'New for VB6.0 starting 1/18/01
'   FormatControls Syntax in every form
'    Dim b As Byte
'    b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())

Public Function AutoFormatControls(frm, TKeyPress() As EsiKeyBd, TGotFocus() As EsiKeyBd, TKeyDown() As EsiKeyBd) As Byte
   ' //Need the following in case of a untrapped Control Array.
   ' Manual Code those from Module Procedures.
   Dim bByte As Byte
   Dim i As Integer
   Dim a As Integer
   Dim b As Integer
   Dim C As Integer
   Dim n As Integer
   
   Dim ESI_txtKeyPress() As New EsiKeyBd
   Dim ESI_txtGotFocus() As New EsiKeyBd
   Dim ESI_txtKeyDown() As New EsiKeyBd
   
   a = -1
   b = -1
   'Have to allow for arrays, etc-blast thru
   On Error Resume Next
   For i = 0 To frm.Controls.Count - 1
      'Part of an Array or label (z1(n))?
      C = frm.Controls(i).Index
      If TypeOf frm.Controls(i) Is SSRibbon Then
         If frm.Controls(i).Name = "ShowPrinters" Then
            Set ESI_cmdShowPrint.esCmdClick = frm.Controls(i)
         End If
      End If
      If Err > 0 And (TypeOf frm.Controls(i) Is TextBox Or _
                      TypeOf frm.Controls(i) Is ComboBox Or TypeOf frm.Controls(i) Is MaskEdBox) Then
         Err = 0
         a = a + 1
         ReDim Preserve ESI_txtKeyPress(a) As New EsiKeyBd
         If frm.Controls(i).Tag <> "9" Then
            b = b + 1
            ReDim Preserve ESI_txtGotFocus(b) As New EsiKeyBd
            ReDim Preserve ESI_txtKeyDown(b) As New EsiKeyBd
         End If
         If TypeOf frm.Controls(i) Is MaskEdBox Then
            Set ESI_txtGotFocus(b).esMskGotFocus = frm.Controls(i)
            Set ESI_txtKeyDown(b).esMskKeyDown = frm.Controls(i)
            Set ESI_txtKeyPress(a).esMskKeyValue = frm.Controls(i)
         End If
         If TypeOf frm.Controls(i) Is TextBox Then
            bByte = True
            Select Case Val(frm.Controls(i).Tag)
               Case 1
                  Set ESI_txtKeyPress(a).esTxtKeyValue = frm.Controls(i)
               Case 3
                  Set ESI_txtKeyPress(a).esTxtKeyCase = frm.Controls(i)
               Case 4
                  Set ESI_txtKeyPress(a).esTxtKeyDate = frm.Controls(i)
               Case 5
                  Set ESI_txtKeyPress(a).esTxtKeyTime = frm.Controls(i)
               Case 9
                  Set ESI_txtKeyPress(a).esTxtKeyMemo = frm.Controls(i)
                  bByte = False
               Case Else
                  Set ESI_txtKeyPress(a).esTxtKeyCheck = frm.Controls(i)
            End Select
            If bByte Then
               Set ESI_txtGotFocus(b).esTxtGotFocus = frm.Controls(i)
               Set ESI_txtKeyDown(b).estxtKeyDown = frm.Controls(i)
            End If
         Else
            If TypeOf frm.Controls(i) Is ComboBox Then
               Set ESI_txtGotFocus(b).esCmbGotfocus = frm.Controls(i)
               Select Case Val(frm.Controls(i).Tag)
                  Case 1
                     Set ESI_txtKeyPress(a).esCmbKeyValue = frm.Controls(i)
                  Case 4
                     Set ESI_txtKeyPress(a).esCmbKeyDate = frm.Controls(i)
                  Case 8
                     Set ESI_txtKeyPress(a).esCmbKeylock = frm.Controls(i)
                     frm.Controls(i).ForeColor = ES_BLUE
                  Case Else
                     Set ESI_txtKeyPress(a).esCmbKeyCase = frm.Controls(i)
               End Select
            End If
         End If
      End If
   Next
   TGotFocus() = ESI_txtGotFocus()
   TKeyPress() = ESI_txtKeyPress()
   TKeyDown() = ESI_txtKeyDown()
   Erase ESI_txtKeyPress()
   Erase ESI_txtGotFocus()
   Erase ESI_txtKeyDown()
   AutoFormatControls = 1
   
End Function

'New 2/15/01 Adds sort to Activity-INNUMBER in InvaTable

Public Function GetLastActivity() As Long
   Dim rdoAct As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT MAX(INNUMBER) FROM InvaTable"
   bSqlRows = GetDataSet(rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         If Not IsNull(.rdoColumns(0)) Then
            GetLastActivity = .rdoColumns(0)
         Else
            GetLastActivity = 0
         End If
         .Cancel
      End With
   End If
   Exit Function
modErr1:
   GetLastActivity = 0
   
End Function

'2/28/01 correct Inventory

Public Sub AdjustInventory(sPassedPart As String)
   Dim rdoAct As rdoResultset
   Dim cQuantity As Currency
   On Error Resume Next
   sPassedPart = Compress(sPassedPart)
   sSql = "SELECT SUM(INAQTY) FROM InvaTable WHERE INPART='" _
          & sPassedPart & "' "
   Set rdoAct = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not IsNull(rdoAct.rdoColumns(0)) Then _
                 cQuantity = rdoAct.rdoColumns(0)
   sSql = "UPDATE PartTable SET PAQOH=" & cQuantity _
          & " WHERE PARTREF='" & sPassedPart & "' "
   RdoCon.Execute sSql, rdExecDirect
   DoEvents
   Set rdoAct = Nothing
   
End Sub

Public Function IllegalCharacters(TestNo As Variant) As Byte
   Dim a As Integer
   Dim k As Integer
   Dim sString As String
   
   On Error GoTo modErr1
   sString = Trim$(TestNo)
   a = Len(sString)
   IllegalCharacters = 0
   If a > 0 Then
      For k = 1 To a
         If Mid$(sString, k, 1) = Chr$(33) Or Mid$(sString, k, 1) = Chr$(34) _
                   Or Mid$(sString, k, 1) = Chr$(35) Or Mid$(sString, k, 1) = Chr$(36) _
                   Or Mid$(sString, k, 1) = Chr$(42) Or Mid$(sString, k, 1) = Chr$(44) _
                   Or Mid$(sString, k, 1) = Chr$(38) Or Mid$(sString, k, 1) = Chr$(58) _
                   Or Mid$(sString, k, 1) = Chr$(59) Or Mid$(sString, k, 1) = Chr$(64) _
                   Or Mid$(sString, k, 1) = Chr$(47) Then
            IllegalCharacters = Asc(Mid$(sString, k, 1))
            Exit For
         End If
      Next
   End If
   Exit Function
   
modErr1:
   Resume Moderr2
Moderr2:
   On Error Resume Next
   
End Function

'Retrieves a unique value for a lot number
'Sends a string that, when converted, will
'yield a Datetime stamp value
'The Increment value adds small bits of time
'to avoid caching for multiple receipts
'Have to force cycling lots with the bNextLot variable
'The Random function helps create the lot
'
'21-37315-812530 first is random, second is date, third is time
'An accurate time stamp can be used by converting to 37315.812530
'Last 3/12/02

Public Function GetNextLotNumber() As String
   Dim i As Integer
   Dim l As Long
   Dim s As Double
   Dim sTime As String
   
   On Error Resume Next
   Randomize bNextLot
   i = Int((99 * Rnd) + 1)
   bNextLot = bNextLot + 1
   s = TimeValue(Format(Time, "hh:nn:ss"))
   sTime = sTime & "-" & Format$(s, ".000000") & Format$(bNextLot, "00")
   sTime = Right$(sTime, 6)
   s = DateValue(Format(Now, "mm/dd/yy"))
   sTime = Format$(s, "00000") & "-" & sTime
   sTime = sTime & "-" & Format$(Trim$(Str$(i)), "00")
   GetNextLotNumber = Trim(sTime)
   If bNextLot > 99 Then bNextLot = 0
   For l = 0 To 128000
   Next 'Give some time with using sleep
   
End Function

Public Function GetSystemCaption() As String
   GetSystemCaption = "ES/" & Format$(Now, "yyyy") & "ERP"
   
End Function

Public Sub GetCurrentDatabase()
   'Database
   sDataBase = GetSetting("Esi2000", "System", "CurDatabase", sDataBase)
   If Trim(sDataBase = "") Then sDataBase = "Esi2000Db"
   If UCase$(Left(sDataBase, 6)) = "TESTDB" Then
      bTestDb = 1
      MdiSect.SystemMsg.ForeColor = ES_RED
      MdiSect.tmr4.Enabled = True
      MdiSect.SystemMsg = "*** Warning - Test Database ***"
   Else
      bTestDb = 0
   End If
   
End Sub


'See if the Lots are registered
'Change later to see if Lot Tracking is active
'3/12/02

Public Function CheckLotTracking() As Byte
   Dim RdoLots As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT LOTNUMBER FROM LohdTable WHERE LOTNUMBER=''"
   bSqlRows = GetDataSet(RdoLots, ES_FORWARD)
   If bSqlRows Then RdoLots.Cancel
   CheckLotTracking = 1
   Set RdoLots = Nothing
   Exit Function
   
modErr1:
   On Error GoTo 0
   CheckLotTracking = 0
   
End Function

'See if lots are selected 3/13/02

Public Function CheckLotsActive() As Byte
   Dim RdoLots As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT COLOTSACTIVE FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = GetDataSet(RdoLots, ES_FORWARD)
   If bSqlRows Then
      CheckLotsActive = RdoLots!COLOTSACTIVE
      RdoLots.Cancel
   End If
   Set RdoLots = Nothing
   Exit Function
modErr1:
   CheckLotsActive = 0
   
End Function

'3/26/02 FIFI or LIFO - Default is FIFO

Public Function GetInventoryMethod() As Byte
   Dim RdoInm As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT COLOTSFIFO FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = GetDataSet(RdoInm, ES_FORWARD)
   If bSqlRows Then
      With RdoInm
         If Not IsNull(!COLOTSFIFO) Then
            GetInventoryMethod = !COLOTSFIFO
         Else
            GetInventoryMethod = 1
         End If
         .Cancel
      End With
   End If
   Set RdoInm = Nothing
   Exit Function
   
modErr1:
   GetInventoryMethod = 1
   
End Function

'3/28/02 - Retrieves the next lot record for Lot Tracking

Public Function GetNextLotRecord(sCurrentLot As String) As Long
   Dim RdoLor As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT MAX(LOIRECORD) FROM LoitTable WHERE " _
          & "LOINUMBER='" & sCurrentLot & "'"
   bSqlRows = GetDataSet(RdoLor, ES_FORWARD)
   If bSqlRows Then
      With RdoLor
         If Not IsNull(.rdoColumns(0)) Then
            GetNextLotRecord = .rdoColumns(0) + 1
         Else
            GetNextLotRecord = 2
         End If
         .Cancel
      End With
   Else
      GetNextLotRecord = 2
   End If
   Set RdoLor = Nothing
   Exit Function
   
modErr1:
   GetNextLotRecord = 2
   
End Function
