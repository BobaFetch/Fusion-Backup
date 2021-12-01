Attribute VB_Name = "ESIPROJ"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit

Public bBold As Byte
Public sProcName As String
Public sReportPath As String
Public bNoCrystal As Boolean
Public iZoomLevel As Integer
Public bUserAction As Boolean
Public iBarOnTop As Byte
Public calEx As Boolean


'Cost Constants
Public Const ES_AVERAGECOST As Byte = 1
Public Const ES_STANDARDCOST As Byte = 2

'StrCase funtion contstants
Public Const ES_FIRSTWORD As Byte = 1

'Cursor types
Public Const ES_FORWARD = 0 'Default
Public Const ES_KEYSET = 1
Public Const ES_DYNAMIC = 2
Public Const ES_STATIC = 3

Public Const SO_NUM_FORMAT = "000000"
Public Const SO_NUM_SIZE = 6

' ADO
Public clsADOCon As ClassFusionADO

'Lot Handling
Public Es_TotalLots As Integer
Public Es_LotsSelected As Byte
Public Es_LotSelectionCanceled As Boolean
'6/16/06

Type LotsAvailable
   LotSysId As String
   LotUserId As String
   LotPartRef As String
   LotADate As String
   LotRemQty As Currency
   LotCost As Currency
   LotSelQty As Currency
   LotExpirationDate As String
End Type

Public lots() As LotsAvailable

'Project variables
Public bDataHasChanged As Boolean
Public bAutoCaps As Byte
Public iAutoTips As Byte
'Public iBarOnTop     As Byte
'Public bBold         As Byte
Public bSysCalendar As Byte
Public bEnterAsTab As Byte
Public bRightArrowAsTab As Byte
Public bSysHelp As Byte
Public bSqlRows As Boolean
'public bNextLot As Byte         'Cycle lots
Public bOpenLastForm As Byte
Public bResize As Byte
Public bActiveTab(8) As Byte
Public bInsertOn As Boolean
Public lScreenWidth As Long
Public sCustomReport As String
Public sFilePath As String
Public sHelpType As String
Public sInitials As String
Public sSql As String


Public sAppTitles(8) As String 'Get and store App.Title
Public ES_SYSDATE As Variant 'Server Date/Time to reduce calls
Public ESI_cmdShowPrint As New EsiKeyBd 'Only one per form
'Key trap for Insert
Public Es_frmKeyDown(20) As New EsiKeyBd
Public sRegistryAppTitle As String


Public Const BRACKET_ALL = "<ALL>"

'Test Resolution (only concern is the width for now)
Public Function ScreenResolution() As String
   Dim iWidth As Integer, iHeight As Integer
   iWidth = Screen.Width \ Screen.TwipsPerPixelX
   iHeight = Screen.Height \ Screen.TwipsPerPixelY
   ScreenResolution = iWidth & " X " & iHeight
End Function

Public Sub GetAppTitles()
   sAppTitles(0) = GetSetting("Esi2000", "AppTitle", "admn", sAppTitles(0))
   If sAppTitles(0) = "" Then sAppTitles(0) = "ESI Administration"
   
   sAppTitles(1) = GetSetting("Esi2000", "AppTitle", "sale", sAppTitles(1))
   If sAppTitles(1) = "" Then sAppTitles(1) = "ESI Sales"
   
   sAppTitles(2) = GetSetting("Esi2000", "AppTitle", "engr", sAppTitles(2))
   If sAppTitles(2) = "" Then sAppTitles(2) = "ESI Engineering"
   
   sAppTitles(3) = GetSetting("Esi2000", "AppTitle", "prod", sAppTitles(3))
   If sAppTitles(3) = "" Then sAppTitles(3) = "ESI Production"
   
   sAppTitles(4) = GetSetting("Esi2000", "AppTitle", "invc", sAppTitles(4))
   If sAppTitles(4) = "" Then sAppTitles(4) = "ESI Inventory"
   
   sAppTitles(5) = GetSetting("Esi2000", "AppTitle", "qual", sAppTitles(5))
   If sAppTitles(5) = "" Then sAppTitles(5) = "ESI Quality"
   
   sAppTitles(6) = GetSetting("Esi2000", "AppTitle", "fina", sAppTitles(6))
   If sAppTitles(6) = "" Then sAppTitles(6) = "ESI Finance"
   
   sAppTitles(7) = GetSetting("Esi2000", "AppTitle", "time", sAppTitles(6))
   If sAppTitles(7) = "" Then sAppTitles(6) = "ESI Time"
End Sub

Public Function GetRegistryAppTitle()
    GetRegistryAppTitle = sRegistryAppTitle
End Function

Public Function SetRegistryAppTitle(ByVal sRegAppTitle As String)
    sRegistryAppTitle = sRegAppTitle
End Function

'Copies of this code can used for any key by getting the
'VK (virtual key) in the API

Public Sub ToggleInsertKey(TurnOn As Boolean)
   'To turn Insert on, set turnon to true
   'To turn Insert off, set turnon to false
   Dim bytKeys(255) As Byte
   Dim bInsertKeyOn As Boolean
   Dim lRet As Long
   Dim typOS As OSVERSIONINFO
   MdiSect.OvrPanel.Enabled = False
   
'   typOS.dwOSVersionInfoSize = Len(typOS)
'   lRet = GetVersionEx(typOS)
   'Get status of the 256 virtual keys
'   GetKeyboardState bytKeys(0)
   
'   bInsertKeyOn = bytKeys(VK_INSERT)
'   If bInsertKeyOn <> TurnOn Then 'if current state <> requested state
'      'Get OS, it matters
'      If typOS.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
'         'Win95/98
'         bytKeys(VK_INSERT) = 1
'         SetKeyboardState bytKeys(0)
'      Else
'         'WinNT/2000/XP
'         'Simulate Key Press
'         keybd_event VK_INSERT, &H45, KEYEVENTF_EXTENDEDKEY _
'            Or 0, 0
'         'Simulate Key Release
'         keybd_event VK_INSERT, &H45, KEYEVENTF_EXTENDEDKEY _
'            Or KEYEVENTF_KEYUP, 0
'      End If
'   End If
   
   If TurnOn Then
      bInsertOn = True
      MdiSect.OvrPanel = "INSERT"
      MdiSect.OvrPanel.ToolTipText = "Insert Text Is On (Click me)"
   Else
      bInsertOn = False
      MdiSect.OvrPanel = "OVER"
      MdiSect.OvrPanel.ToolTipText = "Overtype Text Is On (Click me) "
   End If
   SaveSetting "Esi2000", "mngr", "InsertState", Abs(bInsertOn)
   MdiSect.OvrPanel.Enabled = True
   Exit Sub

modErr1:
   MdiSect.OvrPanel.Enabled = True
   On Error GoTo 0
End Sub

Public Function ReplaceString(ByVal OldString As String) As String
   Dim NewString As String
   'Quotation with alternate
   
   'this extra character will exceed character limit (3)
   'NewString = Replace(OldString, Chr$(34), Chr$(146)) '& Chr$(146))   'leave double-quote
   'Apostrophe with alternate
   NewString = Replace(OldString, Chr$(39), Chr$(146))
   ReplaceString = NewString
   
End Function

Public Function GetCurrentPart(PartNumber As String, ByRef PartDescription As Control, _
                            Optional ShowNone As Boolean) As String
   Dim ADOCur As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_GetINVCfindPart '" & Compress(PartNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOCur, ES_FORWARD)
   If bSqlRows Then
      With ADOCur
         GetCurrentPart = "" & Trim(!PartNum)
         PartDescription = "" & Trim(!PADESC)
         ClearResultSet ADOCur
      End With
      bFoundPart = 1
   Else
      If ShowNone Then GetCurrentPart = "NONE" Else _
                                        GetCurrentPart = ""
      PartDescription = "*** Part Number Wasn't Found ***"
      bFoundPart = 0
   End If
   Set ADOCur = Nothing
   Exit Function

modErr1:
   GetCurrentPart = ""
   PartDescription = "*** Part Number Wasn't Found ***"
   bFoundPart = 0

End Function

Sub CloseForms()
   On Error Resume Next
   Dim bFrmCnt As Integer
   Dim i As Integer
   bUserAction = True
   bFrmCnt = 1 '(Forms.Count - 1) ' don't include the MDI select
   
'   For i = 1 To bFrmCnt
'    If (Forms(i).Name <> "CRViewerFrm") Then
'       Unload Forms(i)
'    End If
'   Next
   ' if CR11 form don't unload the form
   ' increment the form item as we unload the form count decreases.
   Do While Forms.count > bFrmCnt
    If (Forms(bFrmCnt).Name <> "CRViewerFrm") Then
        Unload Forms(bFrmCnt)
    Else
        ' increment the item index to next
        bFrmCnt = bFrmCnt + 1
    End If
   Loop

End Sub

'Use Windows messaging to fill Combo Strings 8/17/00
'AddComboStr cmbVnd.hWnd, sString

Public Sub AddComboStr(lhWnd As Long, sString As String)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, ByVal "" & Trim(sString)

End Sub

'Validate Edits after attemped Update of KeySet
'The operation has falled because a Column in the set
'Was changed somewhere else
'Syntax:   If Err > 0 Then validateedit
'Call after .Update command in a Cursor Edit
'2/17/00 cjs
'12/26/05 frm As Form left for backward compatability

Public Sub ValidateEdit(Optional frm As Form)
   Static bByte As Byte
   
   'Clear it if no error. called from FormLoad to reset Static
   'otherwise use example syntax
   
   If Err = 40026 Then Exit Sub 'The Cursor was not opened (usually AddItem)
   If Err = 0 Then bByte = 3
   'Show the error once and again if the continue on
   If bByte = 0 Then
      'This is a 40002 SQL Server subset
      If Left(Err.Description, 5) = "01S03" Then
         MsgBox "The Data Edited Has Been Changed By Another" & vbCrLf _
            & "Process. You Should Reselect The Data And" & vbCrLf _
            & "Refresh The Information.", _
            vbInformation, MdiSect.ActiveForm.Caption
      Else
         'Process all others
         If Err.Number = 40011 Then
            MsgBox "The Data Has Timed Out. This May Be " _
               & "Normal. Close This Form And Reopen.", _
               vbInformation, MdiSect.ActiveForm.Caption
         Else
            If Err.Number <> 40060 Then
               sProcName = "ValidateEdit"
               CurrError.Number = Err.Number
               CurrError.Description = Err.Description
               DoModuleErrors MdiSect.ActiveForm
            End If
         End If
      End If
   End If
   bByte = bByte + 1
   If bByte > 3 Then bByte = 0
   Err.Clear

End Sub

Public Sub GetFavorites(sSection As String)
   Dim iList As Integer
   
   
   For iList = 1 To 11
      sFavorites(iList) = GetSetting("Esi2000", sSection, "Favorite" & Trim(str(iList)), sFavorites(iList))
   Next
   sFavorites(iList) = GetSetting("Esi2000", sSection, "Favorite" & Trim(str(iList)), sFavorites(iList))
   For iList = 1 To 11
      If sFavorites(iList) <> "" Then
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Caption = sFavorites(iList)
      End If
   Next
   If sFavorites(iList) <> "" Then
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Visible = True
      MdiSect.ActiveBar1.Bands("mnuFavorites").Tools("Favor" & Trim(str(iList))).Caption = sFavorites(iList)
   End If
   
   iBarOnTop = GetSetting("Esi2000", "Programs", "BarOnTop", iBarOnTop)
   iAutoTips = GetSetting("Esi2000", "Programs", "AutoTipsOn", iAutoTips)
   If iAutoTips = 1 Then
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips On"
   Else
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorTips").Caption = "Auto Tips Off"
   End If
   ShowHideTopBar
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", bEnterAsTab)
   bRightArrowAsTab = GetSetting("Esi2000", "System", "RightArrowAsTab", bRightArrowAsTab)
   If RunningInIDE Then
      sReportPath = GetSetting("Esi2000", "System", "ReportPath", sReportPath)
   End If
   If sReportPath = "" Then sReportPath = App.Path & "\"
   bResize = GetSetting("Esi2000", "System", "ResizeForm", bResize)
   GetCrystalZoom

End Sub

Public Sub FillVendors(Optional frm As Form)
   Dim ADOVed As ADODB.Recordset
   On Error GoTo modErr1
   If frm Is Nothing Then
      Set frm = MdiSect.ActiveForm
   End If
   sSql = "Qry_FillSortedVendors"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOVed, ES_FORWARD)
   If bSqlRows Then
      'On Error Resume Next
      With ADOVed
         frm.cmbVnd = "" & Trim(!VENICKNAME)
         'MDISect.ActiveForm.cmbVnd = "" & Trim(!VENICKNAME)
         
         'the following 2 fields may not exist
         On Error Resume Next
         '            MDISect.ActiveForm.txtNme = "" & Trim(!VEBNAME)
         '            MDISect.ActiveForm.lblNme = "" & Trim(!VEBNAME)
         frm.txtNme = "" & Trim(!VEBNAME)
         frm.lblNme = "" & Trim(!VEBNAME)
         On Error GoTo modErr1
         
         Do Until .EOF
            If Trim(!VENICKNAME) <> "NONE" Then _
                    frm.cmbVnd.AddItem "" & Trim(!VENICKNAME)
            'AddComboStr MdiSect.ActiveForm.cmbVnd.hWnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
         ClearResultSet ADOVed
      End With
   End If
   Set ADOVed = Nothing
   Exit Sub

modErr1:
   sProcName = "FillVendors"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm

End Sub

Sub CloseFiles()
   On Error Resume Next
   MdiSect.Cdi.HelpCommand = cdlHelpQuit
   Close
   'clsadocon.Close
   InvalidateRect 0&, 0&, False
   'Set RdoEnv = Nothing
   Set MdiSect = Nothing
   End
   
End Sub

Public Sub CheckKeys(KeyCode As Integer)
   'use in KeyDown
   'not for combo boxes or memo fields
   'to use vbKeyinsert you must have a label
   'name InsPanel (or something else)
   bUserAction = True
   If KeyCode = vbKeyDown Then
      keybd_event VK_TAB, 0, 0, 0
   Else
      If KeyCode = vbKeyUp Then
         SendKeys "+{TAB}"
      End If
   End If

End Sub


'Use constants to return cost type
'Public Const ES_AVERAGECOST As Byte = 1
'Public Const ES_STANDARDCOST As Byte = 2
'9/28/04 Optioned CostType and default to Standard Cost

Public Function GetPartCost(PartRef As String, Optional CostType As Byte) _
                         As Currency
   Dim CostRes As ADODB.Recordset
   
   If CostType = 0 Then CostType = 2
   On Error GoTo modErr1
   PartRef = Compress(PartRef)
   sSql = "Qry_GetPartCost '" & PartRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, CostRes)
   If bSqlRows Then
      With CostRes
         If CostType = ES_AVERAGECOST Then
            GetPartCost = Format(!PAAVGCOST, ES_QuantityDataFormat)
         Else
            GetPartCost = Format(!PASTDCOST, ES_QuantityDataFormat)
         End If
         ClearResultSet CostRes
      End With
   Else
      GetPartCost = 0
   End If
   Set CostRes = Nothing
   Exit Function

modErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   On Error GoTo 0
   GetPartCost = 0

End Function



Public Sub KeyCase(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   bUserAction = True
   If bRightArrowAsTab And KeyAscii = 39 Then
     If MdiSect.ActiveForm.ActiveControl.SelLength = Len(MdiSect.ActiveForm.ActiveControl) Then
        KeyAscii = 0
        keybd_event VK_TAB, 0, 0, 0
        Exit Sub
    End If
   End If
   If KeyAscii = vbKeyReturn Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
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
   If bRightArrowAsTab And KeyAscii = 39 And MdiSect.ActiveForm.ActiveControl.SelLength = Len(MdiSect.ActiveForm.ActiveControl) Then
        KeyAscii = 0
        keybd_event VK_TAB, 0, 0, 0
        Exit Sub
   End If

   If KeyAscii = vbKeyReturn Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
      End If
   Else
      On Error Resume Next
      If Not bInsertOn Then
         If Len(MdiSect.ActiveForm.ActiveControl) > 0 Then _
                If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"
      End If
   End If
   
End Sub

Public Sub KeyDate(KeyAscii As Integer)
   'Changes ".", " " and "-" to "/" for dates
   'syntax in Keypress: KeyDate KeyAscii
   If KeyAscii = vbKeyReturn Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
      End If
   Else
      '3/20/07 Revised
      If KeyAscii < 10 Then Exit Sub
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
   If KeyAscii = vbKeyReturn Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
         Exit Sub
      End If
   Else
      If Not bInsertOn Then
         If KeyAscii > 13 Then
            'SendKeys "{DEL}"       'why delete the next character?
         End If
      End If
      If KeyAscii = 8 Or KeyAscii = 9 Then Exit Sub
      If KeyAscii = 65 Then KeyAscii = 97
      If KeyAscii = 80 Then KeyAscii = 112
      If KeyAscii = 45 Then KeyAscii = 58
      If KeyAscii = 46 Then KeyAscii = 58          '. -> :
      If KeyAscii = 58 Or KeyAscii = 97 Or KeyAscii = 112 Then
         Exit Sub
      End If
      Select Case KeyAscii
         Case 43
            KeyAscii = 112
         Case Is < 43, 47, Is > 57
            KeyAscii = 0
      End Select
   End If
   
End Sub

Public Sub KeyValue(KeyAscii)
   'Allows only numbers, "-" and "." for value
   'fields like money or quantities
   'syntax in Keypress: KeyValue KeyAscii
   'Debug.Print "esiproj KeyValue numbers only"
   bUserAction = True
   If KeyAscii = vbKeyReturn Then
      If bEnterAsTab Then
         KeyAscii = 0
         keybd_event VK_TAB, 0, 0, 0
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


'Changed 1/5/04 to be more consistant with design of
'Windows NT, Windows 2000 and WindowsXP
'Was Screen.MousePointer/MCursor as integer

Sub MouseCursor(MCursor As Byte)
   MdiSect.MousePointer = MCursor
   bUserAction = True
End Sub



'Public Function ParseComment(TestCmt As Variant, Optional ParseLineFeed As Boolean) As String
'   'Replace double quotes for SQL Server text fields
'   Dim A As Integer
'   Dim d As Integer
'   Dim E As Integer
'   Dim g As Integer
'   Dim K As Integer
'   Dim n As Integer
'   Dim NewComment As String
'
'   'WARNING: THIS IS TOXIC!  It REPLACES INCHES WITH FEET, ETC.  Use SQLString in the SSQL = statement instead.
'   On Error GoTo modErr1
'   NewComment = RTrim(TestCmt)
'   A = Len(NewComment)
'   If ParseLineFeed Then _
'      NewComment = Replace(NewComment, vbCrLf, " ")
'   NewComment = NewComment & Chr$(255)
'   K = 1
'   Do Until K > A
'      If Mid(NewComment, K, 1) = Chr(34) Then
'         Mid(NewComment, K, 1) = Chr(39)
'         NewComment = NewComment & Chr$(255)
'         E = Len(NewComment)
'         For g = E To K + 1 Step -1
'            Mid(NewComment, g, 1) = Mid(NewComment, g - 1, 1)
'         Next
'         Mid(NewComment, K + 1, 1) = Chr(39)
'         K = K - 1
'      End If
'      K = K + 1
'   Loop
'
'   NewComment = RTrim$(NewComment)
'   NewComment = Replace(NewComment, Chr(255), "")
'   ParseComment = NewComment
'   Exit Function
'
'modErr1:
'   Resume modErr2
'modErr2:
'   On Error GoTo 0
'
'End Function

'Used to set an MDIChild form position and provide ToolTips
'See ES_LIST, ES_DONTLIST Constants
'See ES_RESIZE, ES_DONTRESIZE Constansts
'2/5/04 Added trap for listing
'8/2/04 Added LockWindowUpdate
'4/7/05 trimmed "ALL" in Currents
'6/30/05 Added form.HelpContextID
'7/25/05 vContextID/GetSetting to recall Topic ID
'2/7/07 Was FormLoad

Sub FormLoad(frm As Form, Optional DontList As Boolean, Optional noResize As Boolean)
   Dim iList As Integer
   Dim vContextID As Variant
   
   bDataHasChanged = False
   
   If Trim(cUR.CurrentCustomer) = "ALL" Then cUR.CurrentCustomer = ""
   If Trim(cUR.CurrentPart) = "ALL" Then cUR.CurrentPart = ""
   If Trim(cUR.CurrentVendor) = "ALL" Then cUR.CurrentVendor = ""
   If Not noResize Then SetFormSize frm
   
   frm.Move 0, 0
   If UCase$(Left$(Forms(Forms.count - 1).Name, 3)) = "zGr" Then
      On Error Resume Next
      'frm.Tab1.TabHeight = 300
      frm.Tag = "TAB"
      iList = 1
      DontList = True
   Else
      frm.KeyPreview = True
      Set Es_frmKeyDown(Forms.count - 1).esFormKeyDown = frm
      iList = Forms.count - 1
      If frm.Name <> "PurcPRe01a" Then _
         If iList > 2 Then DontList = True
   End If
   If Not DontList Then sCurrForm = frm.Caption
   
   'Help (OpenHelpContext)
   '    Default F1 Help for report dialogs
   '    Finance is different
   If UCase$(Left$(sProgName, 3)) = "FIN" Then
      If UCase$(Mid$(frm.Name, 6, 1)) = "P" Then frm.HelpContextID = 907
   Else
      If UCase$(Mid$(frm.Name, 4, 1)) = "P" Then
         If Val(Right$(frm.Name, 2)) > 0 Then frm.HelpContextID = 907
      End If
   End If
   'Recall saved TopicID
   vContextID = GetSetting("Esi2000", "Help", frm.Caption, vContextID)
   If Val(vContextID) <> 0 Then frm.HelpContextID = vContextID
   If frm.Tag = "TAB" Then frm.HelpContextID = 923
   If frm.HelpContextID = 0 Then frm.HelpContextID = 904
   'End Help
   bDataHasChanged = False
   bUserAction = True
   'bNextLot = 0
   Err.Clear
   
   On Error Resume Next
   sProcName = ""
   frm.cmdCan.Cancel = True
   frm.KeyPreview = True
   ES_SYSDATE = GetServerDateTime()
   If Not DontList Then iList = SetRecent(frm)
   GetCurrentSelections
   LockWindowUpdate 0

End Sub


Public Function SetRecent(frm As Form) As Integer
   Dim a As Integer
   Dim iList As Integer
   Static iListcount As Integer
   Dim sTemp As String
   
   On Error GoTo modErr1
   If iListcount < 50 Then
      For iList = iListcount To 0 Step -1
         sSession(iList + 1) = sSession(iList)
      Next
      iListcount = iListcount + 1
      sSession(0) = frm.Caption
   End If
   
   Erase sRecent
   a = 0
   For iList = 1 To 5
      sTemp = MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList - 1))).Caption
      If sTemp = frm.Caption Then sTemp = ""
      If Len(Trim(sTemp)) < 3 Then sTemp = ""
      If sTemp <> "" Then
         a = a + 1
         sRecent(a) = sTemp
      End If
   Next
   If a > 4 Then a = 4
   sRecent(0) = frm.Caption
   For iList = 0 To 4
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = False
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Caption = Trim(str(iList))
   Next
   For iList = 0 To a
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = True
      MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Caption = sRecent(iList)
   Next
   Exit Function

modErr1:
   Resume modErr2
modErr2:
   'Ignore them
   On Error Resume Next

End Function



'Small message instead of a MsgBox
'Timer On either True or False
'Syntax: SysMsg "User Message", True
'SysMessage may be up to 24 characters

Sub SysMsg(SysMessage As String, TimerOn As Byte, Optional frm As Form)
   On Error GoTo modErr1
   SysMsgBox.tmr1.Enabled = TimerOn
   SysMsgBox.msg = SysMessage
   Beep
   SysMsgBox.Show vbModal
   On Error Resume Next
   frm.Refresh
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   'Can't show modal form on MdiChildren
   On Error Resume Next
   SysMsgBox.tmr1.Enabled = TimerOn
   SysMsgBox.msg = SysMessage
   SysMsgBox.Show
   
End Sub

'Test for a valid date otherwise Use Today
'Syntax:  txtDte = CheckDate(txtDte)
'3/28/05 revised with IsDate
'4/21/05 Replaced function to allow dates like 042005 (by popular demand)

Public Function CheckDate(NewDate As String)
   On Error GoTo modErr1
   If Val(NewDate) > 12 And Len(NewDate) = 6 Then
      NewDate = Left$(NewDate, 2) & "/" & Mid$(NewDate, 3, 2) & "/" & Right$(NewDate, 2)
   End If
   If IsDate(NewDate) Then
      CheckDate = Format(NewDate, "mm/dd/yy")
   Else
      CheckDate = Format(ES_SYSDATE, "mm/dd/yy")
   End If
   Exit Function

modErr1:
   On Error Resume Next
   CheckDate = Format(ES_SYSDATE, "mm/dd/yy")
   
End Function

'Test for a valid date otherwise Use Today
'Syntax:  txtDte = CheckDateYYYY(txtDte)
Public Function CheckDateYYYY(NewDate As String)
   On Error GoTo modErr1
   Dim slash1 As Integer, slash2 As Integer, mo As String, day As String, year As String
   slash1 = InStr(1, NewDate, "/")
   If slash1 > 1 Then
      mo = Left(NewDate, slash1 - 1)
      slash2 = InStr(slash1 + 1, NewDate, "/")
      If slash2 > slash1 + 1 Then
         day = Mid(NewDate, slash1 + 1, slash2 - slash1 - 1)
         year = Mid(NewDate, slash2 + 1)
         If Len(year) = 2 Then
            year = "20" & year
         End If
         NewDate = mo & "/" & day & "/" & year
         If IsDate(NewDate) Then
            CheckDateYYYY = Format(NewDate, "mm/dd/yyyy")
            Exit Function
         End If
      End If
   End If
   
   CheckDateYYYY = Format(ES_SYSDATE, "mm/dd/yyyy")
   Exit Function

modErr1:
   On Error Resume Next
   CheckDateYYYY = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Function


Public Function CheckDateEx(NewDate As String)
   On Error GoTo modErr1
   If Val(NewDate) > 12 And Len(NewDate) = 6 Then
      NewDate = Left$(NewDate, 2) & "/" & Mid$(NewDate, 3, 2) & "/" & Right$(NewDate, 2)
   End If
   If IsDate(NewDate) Then
      CheckDateEx = Format(NewDate, "mm/dd/yyyy")
   Else
      CheckDateEx = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   Exit Function

modErr1:
   On Error Resume Next
   CheckDateEx = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Function

Public Sub KeyMemo(KeyAscii As Integer)
   If Not bInsertOn Then
      If KeyAscii > 32 Then SendKeys "+{RIGHT}{DEL}", True
   End If
   
End Sub

'Adjust Box length to fit data fields
'Also checks to make sure there and no (') to mess SQL Server up

Public Function CheckLen(sTextBox As String, iTextLength As Integer) As String
   sTextBox = Trim(sTextBox)
   If Len(sTextBox) > iTextLength Then sTextBox = Left(sTextBox, iTextLength)
   CheckLen = sTextBox
   iTextLength = InStr(1, CheckLen, Chr$(39))
   If iTextLength > 0 Then CheckLen = ReplaceString(CheckLen)

End Function

Public Function CheckLenOnly(sTextBox As String, iTextLength As Integer) As String
   sTextBox = Trim(sTextBox)
   If Len(sTextBox) > iTextLength Then sTextBox = Left(sTextBox, iTextLength)
   CheckLenOnly = sTextBox
End Function

Public Function ReplaceSingleQuote(ByVal OldString As String) As String
   Dim NewString As String
   'Quotation with alternate
   
   'this extra character will exceed character limit (3)
   NewString = Replace(OldString, Chr$(39), Chr$(39) & Chr$(39))
   ReplaceSingleQuote = NewString
   
End Function

Public Function ReplaceDoubleQuote(ByVal OldString As String) As String
   Dim NewString As String
   'Quotation with alternate
   
   'this extra character will exceed character limit (3)
   NewString = Replace(OldString, Chr$(34), Chr$(32))
   ReplaceDoubleQuote = NewString
   
End Function

Public Sub KeyLock(KeyAscii As Integer)
   'All uppercase
   'syntax in Keypress Procedure KeyCase KeyAscii
   'If Combo <> True Then Combo = False
   bUserAction = True
   If KeyAscii = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
   'If KeyAscii > 9 Then KeyAscii = 0
   '*new
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii > 13 Then SendKeys "+{RIGHT}{DEL}"

End Sub

'Finds and updates Average Cost for a Part

Public Sub AverageCost(sPassedPart As String)
   Dim ActRs As ADODB.Recordset
   Dim cAverageCost As Currency
   
   On Error GoTo whoops
   sPassedPart = Compress(sPassedPart)
   sSql = "SELECT SUM(INAMT*Abs(INAQTY))/SUM(Abs(INAQTY)) " _
      & "From InvaTable WHERE INAQTY<>0 AND " _
      & "(INPART='" & sPassedPart & "') "
   Set ActRs = clsADOCon.GetRecordSet(sSql, ES_STATIC)
      
   'Set ActRs = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not IsNull(ActRs.Fields(0)) Then
      If Val(ActRs.Fields(0)) > 0 Then cAverageCost = ActRs.Fields(0)
   End If
   ClearResultSet ActRs
   
   sSql = "UPDATE PartTable SET " _
      & "PAAVGCOST=" & Format(cAverageCost, "#####.0000") _
      & " WHERE PARTREF='" & sPassedPart & "' "
   clsADOCon.ExecuteSql sSql
   Set ActRs = Nothing
   Exit Sub
   
'modErr1:
'   Resume modErr2
'modErr2:
'   On Error Resume Next
'   Set ActRs = Nothing

whoops:
   sProcName = "AverageCost"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub


'12/26/05 frm as Form left for backward compatability

Public Sub FillCustomers(Optional frm As Form, Optional ShowAll As Boolean)

   If frm Is Nothing Then
      Set frm = MdiSect.ActiveForm
   End If
   Dim combo As ComboBox
   Set combo = frm.cmbCst
   
   MouseCursor 13
   Dim ADOCst As ADODB.Recordset
   On Error GoTo modErr1
   combo.Clear
   
   '    'if showing 'ALL', add it to the list
   '    If ShowALL Then
   '        AddComboStr combo.hWnd, "ALL"
   '    End If
   
   sSql = "Qry_FillCustomerCombo"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOCst, ES_FORWARD)
   If bSqlRows Then
      With ADOCst
         Do Until .EOF
            AddComboStr combo.hwnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet ADOCst
      End With
   End If
   
   Set ADOCst = Nothing
   
   If ShowAll Then
      combo = "ALL"
   End If
   
   MouseCursor 0
   Exit Sub
   
modErr1:
   sProcName = "fillcustomers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set ADOCst = Nothing
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


Public Sub FindCustomer(frm As Form, sCustomerNickname, Optional bNeedsMore As Byte)
   Dim CusRes As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FindCustomer '" & Compress(sCustomerNickname) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, CusRes)
   If bSqlRows Then
      With CusRes
         On Error Resume Next
         frm.lblCst = "" & Trim(!CUNICKNAME)
         frm.cmbCst = "" & Trim(!CUNICKNAME)
         frm.lblNme = "" & Trim(!CUNAME)
         frm.txtNme = "" & Trim(!CUNAME)
         If bNeedsMore Then
            frm.txtDis = Format(!CUDISCOUNT, "#0.00")
            frm.txtFra = Format(!CUFRTALLOW, ES_QuantityDataFormat)
            frm.txtFrd = Format(!CUFRTDAYS, "##0")
         End If
         ClearResultSet CusRes
      End With
   Else
      On Error Resume Next
      frm.lblNme = ""
      frm.txtNme = "*** Customer Wasn't Found ***"
      If Trim(frm.cmbCst) = "" Then frm.txtNme = "*** No Customer Selected ***"
   End If
   Set CusRes = Nothing
   Exit Sub
   
modErr1:
   sProcName = "findcustomer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
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


Public Sub GetCurrentSelections()
   cUR.CurrentPart = GetSetting("Esi2000", "Current", "Part", cUR.CurrentPart)
   cUR.CurrentVendor = GetSetting("Esi2000", "Current", "Vendor", cUR.CurrentVendor)
   cUR.CurrentCustomer = GetSetting("Esi2000", "Current", "Customer", cUR.CurrentCustomer)
   cUR.CurrentRegion = GetSetting("Esi2000", "Current", "Region", cUR.CurrentRegion)
   bEnterAsTab = GetSetting("Esi2000", "System", "EnterAsTab", bEnterAsTab)
   bRightArrowAsTab = GetSetting("Esi2000", "System", "RightArrowAsTab", bRightArrowAsTab)
End Sub

Public Sub SaveCurrentSelections()
   SaveSetting "Esi2000", "Current", "Part", cUR.CurrentPart
   SaveSetting "Esi2000", "Current", "Vendor", cUR.CurrentVendor
   SaveSetting "Esi2000", "Current", "Customer", cUR.CurrentCustomer
   SaveSetting "Esi2000", "Current", "Region", cUR.CurrentRegion

End Sub

'Called From FormLoad
'Sets the form size to allow Resize.ocx to do whatever it does...
'See ES_RESIZE, ES_DONTRESIZE Constansts
'6/14/06 See CommandButton-Changed Apply and ALL captions
'6/14/06 Blocked SSRibbon for optDis/optPrn
'6/23/06 Removed Threed32.ocx references

Public Sub SetFormSize(frm As Form)
   Dim bByte As Byte
   Dim iList As Integer
   Dim iScreenSize As Integer
   Dim cNewSize As Currency
   
   On Error Resume Next
   LockWindowUpdate frm.hwnd
   iScreenSize = Screen.Width \ Screen.TwipsPerPixelX
   If bResize = 1 Then
'      If iScreenSize > 640 Then '640 X 480
'         If iScreenSize < 1024 Then '800 X 600
'            cNewSize = 1.05
'         ElseIf iScreenSize = 1024 Then cNewSize = 1.1 '1024 X 768
'         ElseIf iScreenSize = 1280 Then cNewSize = 1.13 '1280
'         Else
'            If iScreenSize > 1590 Then cNewSize = 1.14 'Even Greater
'         End If
'         frm.Height = frm.Height * cNewSize
'         frm.Width = frm.Width * cNewSize
'      End If
      If iScreenSize > 640 Then '640 X 480
         If iScreenSize < 1024 Then '800 X 600
            cNewSize = 1.05
         ElseIf iScreenSize = 1024 Then cNewSize = 1.1 '1024 X 768
         ElseIf iScreenSize <= 1280 Then cNewSize = 1.13 '1280
         Else
            cNewSize = 1.14
         End If
         frm.Height = frm.Height * cNewSize
         frm.Width = frm.Width * cNewSize
      End If
   Else
      frm.ReSize1.Enabled = False
   End If
   
   'Setup the controls and bring some order to the place. Like it or not
   'colors 10/2/01 - added here 12/05/04
   If frm.BackColor <> ES_ViewBackColor Then frm.BackColor = Es_FormBackColor
   bByte = 0
   For iList = 0 To frm.Controls.count - 1
      If UCase$(frm.Controls(iList).Name) = "PRG1" Then frm.Controls(iList).Visible = False
      If TypeOf frm.Controls(iList) Is TextBox Then
         frm.Controls(iList).BackColor = Es_TextBackColor
         frm.Controls(iList).ForeColor = Es_TextForeColor
         If frm.Controls(iList).Tag <> 5 Then
            frm.Controls(iList).Text = " "
            If Trim(frm.Controls(iList).ToolTipText) = "" Then
               If frm.Controls(iList).Tag = 1 Then
                  frm.Controls(iList).ToolTipText = " Value Formatted TextBox "
               ElseIf frm.Controls(iList).Tag = 2 Then
                  frm.Controls(iList).ToolTipText = " Any Format TextBox "
               ElseIf frm.Controls(iList).Tag = 3 Then
                  frm.Controls(iList).ToolTipText = " UpperCase Formatted TextBox "
               ElseIf frm.Controls(iList).Tag = 4 Then
                  frm.Controls(iList).ToolTipText = " Date Formatted TextBox (09/01/08, 09-01-05 Or 09.01.05)"
               ElseIf frm.Controls(iList).Tag = 5 Then
                  frm.Controls(iList).ToolTipText = " Time Formatted TextBox (01:00a Or 09:30p)"
               End If
            End If
         Else
            frm.Controls(iList).Text = "  :  "
         End If
      ElseIf TypeOf frm.Controls(iList) Is Frame Then
         If frm.Controls(iList).Name = "fraPrn" Then
            frm.Controls(iList).Top = 360
         End If
      ElseIf TypeOf frm.Controls(iList) Is SSRibbon Then
         If frm.Controls(iList).Name = "cmdHlp" Then
            frm.Controls(iList).Visible = False
         ElseIf frm.Controls(iList).Name = "optDis" Then
            frm.Controls(iList).Picture = MdiSect.XDisplay.Picture
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
         ElseIf frm.Controls(iList).Name = "optPrn" Then
            frm.Controls(iList).Picture = MdiSect.XPrinter.Picture
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
        ElseIf frm.Controls(iList).Name = "ShowPrinters" Then
           frm.Controls(iList).Picture = MdiSect.XPrinter_small.Picture
           frm.Controls(iList).PictureUp = MdiSect.XPrinter_small.Picture
           frm.Controls(iList).PictureDn = MdiSect.XPrinter_small.Picture
           frm.Controls(iList).MaskColor = RGB(212, 208, 200)
         End If
      ElseIf TypeOf frm.Controls(iList) Is CommandButton Then
         If frm.Controls(iList).Style = 1 Then
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
            frm.Controls(iList).UseMaskColor = True
         End If
         
         frm.Controls(iList).ForeColor = Es_TextForeColor
         If frm.Controls(iList).Name = "cmdCan" And frm.Controls(iList).Top < 200 Then
            frm.Controls(iList).Top = 0
            frm.Controls(iList).ToolTipText = " Close This Form (esc) "
         End If
         
         If frm.Controls(iList).Name = "cmdHlp" Then
            frm.Controls(iList).DownPicture = MdiSect.XPHelpDn.Picture
            frm.Controls(iList).Picture = MdiSect.XPHelpDn.Picture
            frm.Controls(iList).Height = 375
            frm.Controls(iList).Width = 375
            frm.Controls(iList).Visible = False
         ElseIf frm.Controls(iList).Name = "cmdComments" Then
            frm.Controls(iList).Picture = MdiSect.imgStandardComment.Picture
            frm.Controls(iList).DownPicture = MdiSect.imgStandardComment.Picture
            frm.Controls(iList).DisabledPicture = MdiSect.imgStandardComment.Picture
            
         ElseIf frm.Controls(iList).Name = "cmdVew" Then
            frm.Controls(iList).Picture = MdiSect.imgPartList.Picture
            frm.Controls(iList).DownPicture = MdiSect.imgPartList.Picture
            frm.Controls(iList).DisabledPicture = MdiSect.imgPartList.Picture
         ElseIf frm.Controls(iList).Name = "cmdPrevious" Then
            frm.Controls(iList).Picture = MdiSect.imgPartList.Picture
            frm.Controls(iList).DownPicture = MdiSect.imgPartList.Picture
            frm.Controls(iList).DisabledPicture = MdiSect.imgPartList.Picture

         ElseIf frm.Controls(iList).Name = "cmdPrt" Then
            frm.Controls(iList).Picture = MdiSect.imgNewPart.Picture
            frm.Controls(iList).DownPicture = MdiSect.imgNewPart.Picture
            frm.Controls(iList).DisabledPicture = MdiSect.imgNewPart.Picture
         ElseIf frm.Controls(iList).Name = "cmdFnd" Then
            frm.Controls(iList).Picture = MdiSect.imgPartFind.Picture
            frm.Controls(iList).DownPicture = MdiSect.imgPartFind.Picture
         ElseIf frm.Controls(iList).Name = "optDis" Then
            frm.Controls(iList).Picture = MdiSect.XDisplay.Picture
            frm.Controls(iList).DownPicture = MdiSect.XDisplay.Picture
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
         ElseIf frm.Controls(iList).Name = "optPrn" Then
            frm.Controls(iList).Picture = MdiSect.XPrinter.Picture
            frm.Controls(iList).DownPicture = MdiSect.XPrinter.Picture
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
         ElseIf frm.Controls(iList).Name = "cmdFnd" Then
            frm.Controls(iList).MaskColor = RGB(212, 208, 200)
            frm.Controls(iList).UseMaskColor = True
         Else
            If frm.Controls(iList).Name = "ShowPrinters" Then
               frm.Controls(iList).AutoSize = 2
               frm.Controls(iList).Left = 360
               frm.Controls(iList).Picture = MdiSect.XPrinter.Picture
               frm.Controls(iList).PictureUp = MdiSect.XPrinter.Picture
               frm.Controls(iList).PictureDn = MdiSect.XPPrinterDn.Picture
               frm.Controls(iList).MaskColor = RGB(212, 208, 200)
               frm.lblPrinter.Left = 720
               frm.lblPrinter.Width = 2200
            End If
         End If
      ElseIf TypeOf frm.Controls(iList) Is Frame Then
         frm.Controls(iList).BackColor = ES_SystemBackcolor
      ElseIf TypeOf frm.Controls(iList) Is CheckBox Then
         frm.Controls(iList).BackColor = ES_SystemBackcolor
         If Left$(frm.Controls(iList).Caption, 2) = "__" Then
            frm.Controls(iList).ForeColor = Es_CheckBoxForeColor
         Else
            frm.Controls(iList).ForeColor = Es_TextForeColor
         End If
         If Trim(frm.Controls(iList).ToolTipText) = "" Then _
                 frm.Controls(iList).ToolTipText = " CheckBox - SpaceBar Or Click To Select "
         
      ElseIf TypeOf frm.Controls(iList) Is OptionButton Then
         frm.Controls(iList).BackColor = ES_SystemBackcolor
         frm.Controls(iList).ForeColor = Es_TextForeColor
      Else
         If TypeOf frm.Controls(iList) Is ComboBox Then
            frm.Controls(iList).BackColor = Es_TextBackColor
            frm.Controls(iList).ForeColor = Es_TextForeColor
            If frm.Controls(iList).Tag <> "" Then
               If frm.Controls(iList).Tag = "4" Then
                  frm.Controls(iList).ToolTipText = " Date As 09/01/05, " _
                               & "09-01-05, 09.01.05 Or Pull Down "
               End If
               If Trim(frm.Controls(iList).ToolTipText) = "" Then
                  If frm.Controls(iList).Tag = 1 Then
                     frm.Controls(iList).ToolTipText = " Value Formatted ComboBox "
                  ElseIf frm.Controls(iList).Tag = 2 Then
                     frm.Controls(iList).ToolTipText = " Any Format ComboBox "
                  ElseIf frm.Controls(iList).Tag = 3 Then
                     frm.Controls(iList).ToolTipText = " UpperCase Formatted ComboBox "
                  ElseIf frm.Controls(iList).Tag = 8 Then
                     frm.Controls(iList).ToolTipText = " Data Entry Edit Is Locked "
                  End If
               End If
               
            Else
               If frm.Controls(iList).Name = "txtBeg" Or frm.Controls(iList).Name = "txtEnd" _
                               Or frm.Controls(iList).Name = "txtDte" Then
                  frm.Controls(iList).Tag = 4
                  frm.Controls(iList).ToolTipText = " Date as 09/01/05,09-01-05,09.01.05 Or Pull Down "
               End If
            End If
         End If
      End If
   Next
   'Tool tips
   '4/18/05
   iAutoTips = 1
   If iAutoTips = 1 Then
   For iList = 0 To frm.Controls.count - 1
      If TypeOf frm.Controls(iList) Is TextBox Then
         If frm.Controls(iList).ToolTipText = "" Then _
                         frm.Controls(iList).Text = " "
         If frm.Controls(iList).ToolTipText = "" Then
            Select Case Val(frm.Controls(iList).Tag)
               Case 1
                  frm.Controls(iList).ToolTipText = " Value (number) "
               Case 3
                  frm.Controls(iList).ToolTipText = " Upper Case Entry "
               Case 4
                  frm.Controls(iList).ToolTipText = " Date as 09/01/05,09-01-05,09.01.05 Or Pull Down "
               Case 5
                  frm.Controls(iList).ToolTipText = " Time as 10:32a "
                  frm.Controls(iList).Text = "  :  "
               Case 9
                  frm.Controls(iList).ToolTipText = " Multiple Line Entry "
               Case Else
                  frm.Controls(iList).ToolTipText = " Any Alpha/Numeric Entry "
            End Select
         End If
      Else
         If TypeOf frm.Controls(iList) Is ComboBox Then
            If frm.Controls(iList).ToolTipText = "" And _
                              frm.Controls(iList).Tag <> 4 Then _
                              frm.Controls(iList).ToolTipText = "ComboBox"
            End If
         End If
         If TypeOf frm.Controls(iList) Is CommandButton Then
            If UCase$(frm.Controls(iList).Caption) = "&ALL" Then frm.Controls(iList).Caption = "A&LL"
            If frm.Controls(iList).Name = "cmdOk" Then _
                            frm.Controls(iList).ToolTipText = " Continue Processing "
            If frm.Controls(iList).Name = "cmdCan" And frm.Controls(iList).Top < 200 Then _
                            frm.Controls(iList).ToolTipText = " Close Form (Escape) "
            If frm.Controls(iList).ToolTipText = "" Then
               If frm.Controls(iList).Name = "cmdUpd" Then
                  frm.Controls(iList).ToolTipText = "Update/Apply Data Changes"
               Else
                  frm.Controls(iList).ToolTipText = "Command Button"
               End If
            End If
            If frm.Controls(iList).Name = "optDis" Then
               frm.Controls(iList).ToolTipText = " Display The Report "
            Else
               If frm.Controls(iList).Name = "optPrn" Then _
                               frm.Controls(iList).ToolTipText = " Print The Report "
            End If
         Else
            If TypeOf frm.Controls(iList) Is ListBox Then
               If frm.Controls(iList).ToolTipText = "" Then _
                               frm.Controls(iList).ToolTipText = "ListBox"
            End If
         End If
         If frm.Controls(iList).Name = "fraPrn" Then bByte = 1
      Next
   End If
   If Left(frm.Name, 3) <> "zGr" Then
      For iList = 1 To 8
         bActiveTab(iList) = bByte
      Next
   End If
   LockWindowUpdate 0
   
End Sub



'2/8/06 Corrected to allow for the possibility that the control has been
'disabled
'2/22/07 Revised Tag etc.

Public Sub SelectFormat(frm As Form)
   'Selects all of the text in fixed length TextBoxes
   'and ComboBoxes. Mostly gets rid of the blinking
   'num locks with VB5.0 (SP2) and SendKeys
   On Error Resume Next
   If Len(frm.ActiveControl) = 0 Then frm.ActiveControl = " "
   If frm.ActiveControl.Tag <> 8 Then
      frm.ActiveControl.SelStart = 0
      frm.ActiveControl.SelLength = Len(frm.ActiveControl.Text)
   End If
   
End Sub




'12/26/05 frm as Form left for backward compatability

Public Function IsControlOnForm(sCtrlNme As String, Optional frm As Form) As Boolean
    Dim i As Integer
    Dim frmControl As Control
    
    If frm Is Nothing Then Set frm = MdiSect.ActiveForm
    IsControlOnForm = False
    For Each frmControl In frm.Controls
           If frmControl.Name = sCtrlNme Then
                IsControlOnForm = True
                Exit Function
           End If
    Next frmControl
End Function


'12/26/05 frm as Form left for backward compatability

Public Sub FillDivisions(Optional frm As Form)
   Dim AdoDiv As ADODB.Recordset
   On Error GoTo modErr1
   Set AdoDiv = clsADOCon.GetRecordSet("Qry_FillDivisions", ES_STATIC)
   
   'Set RdoDiv = clsadocon.OpenResultset("Qry_FillDivisions", rdOpenForwardOnly, rdConcurReadOnly)
   If Not AdoDiv.BOF And Not AdoDiv.EOF Then
      With AdoDiv
         Do Until .EOF
            If frm Is Nothing Then
                If IsControlOnForm("cmbDiv") Then
                    AddComboStr MdiSect.ActiveForm.cmbDiv.hwnd, "" & Trim(!DIVREF)
                End If
            Else
                AddComboStr frm.cmbDiv.hwnd, "" & Trim(!DIVREF)
            End If
            .MoveNext
         Loop
         ClearResultSet AdoDiv
      End With
   End If
   On Error Resume Next
   Set AdoDiv = Nothing
   Exit Sub
   
modErr1:
   sProcName = "filldivisions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


'12/26/05 frm as Form left for backward compatability

Public Sub FillRegions(Optional frm As Form)
   Dim ADOReg As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillRegions"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOReg, ES_FORWARD)
   If bSqlRows Then
      With ADOReg
         Do Until .EOF
            If frm Is Nothing Then
                AddComboStr MdiSect.ActiveForm.cmbReg.hwnd, "" & Trim(!REGREF)
            Else
                AddComboStr frm.cmbReg.hwnd, "" & Trim(!REGREF)
            End If
            .MoveNext
         Loop
         ClearResultSet ADOReg
      End With
   End If
   Set ADOReg = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillregions"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub FillStatusCode(ByRef cmbStatCode As ComboBox, Optional frm As Form)
   Dim ADOStatCd As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillStatCode"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOStatCd, ES_FORWARD)
   If bSqlRows Then
      With ADOStatCd
         Do Until .EOF
            If frm Is Nothing Then
                AddComboStr MdiSect.ActiveForm.cmbStatID.hwnd, "" & Trim(!STATUS_REF)
            Else
                AddComboStr cmbStatCode.hwnd, "" & Trim(!STATUS_REF)
            End If
            .MoveNext
         Loop
         ClearResultSet ADOStatCd
      End With
   End If
   Set ADOStatCd = Nothing
   Exit Sub
   
modErr1:
   sProcName = "FillStatusCode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'10/25/04 Removed adding a blank item
'10/26/05 frm As Form left for backward compatability

Public Sub FillProductCodes(Optional frm As Form)
   Dim AdoCde As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillProductCodes"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoCde, ES_FORWARD)
   If bSqlRows Then
      With AdoCde
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbCde.hwnd, "" & Trim(!PCCODE)
            .MoveNext
         Loop
         ClearResultSet AdoCde
      End With
   End If
   Set AdoCde = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillproductcodes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'Generally calls Form.cmbCls
'12/26/05 frm As Form left for backward compatability

Public Sub FillProductClasses(Optional frm As Form)
   Dim ADOCls As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillProductClasses"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOCls, ES_FORWARD)
   If bSqlRows Then
      With ADOCls
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbCls.hwnd, "" & Trim(!CCCODE)
            .MoveNext
         Loop
         ClearResultSet ADOCls
      End With
   End If
   Set ADOCls = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillproductclasses"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'12/26/05 frm as Form (left for backward compatability)

Public Sub FillTerms(Optional frm As Form)
   Dim ADOTrm As ADODB.Recordset
   On Error GoTo modErr1
   
   sSql = "SELECT TRMREF FROM StrmTable ORDER BY TRMREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOTrm, ES_FORWARD)
   If bSqlRows Then
      With ADOTrm
         Do Until .EOF
            If frm Is Nothing Then
                AddComboStr MdiSect.ActiveForm.cmbTrm.hwnd, "" & Trim(!TRMREF)
            Else
                AddComboStr frm.cmbTrm.hwnd, "" & Trim(!TRMREF)
            End If
            .MoveNext
         Loop
         ClearResultSet ADOTrm
      End With
   End If
   Set ADOTrm = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'Load and show the SysCalendar from the combo dropdown
'9/13/04 Added the Val parameter to test where date = "ALL"

'Public Sub ShowCalendar(frm As Form, Optional iAdjust As Integer, Optional Cntl As Control)
Public Sub ShowCalendar(frm As Form, Optional iAdjust As Integer)

   'display date selection calendar
   
   Dim iAdder As Integer
   Dim lLeft As Long
   Dim lTop As Long
   Dim sDate As Date
   
   Dim combo As Control
   Set combo = frm.ActiveControl
   
   
   If IsDate(frm.ActiveControl.Text) Then
      'frm.ActiveControl.AddItem frm.ActiveControl.Text
   Else
      'frm.ActiveControl.AddItem Format(Now, "mm/dd/yy")
      combo.Text = Format(Now, "mm/dd/yy")
   End If
   
   'set form to pass date back to
   Set SysCalendar.FromForm = frm

   'On Error Resume Next
   'See if there is a date in the combo
   If IsDate(frm.ActiveControl.Text) Then
      sDate = Format(frm.ActiveControl.Text, "mm/dd/yy")
   Else
      sDate = Format(ES_SYSDATE, "mm/dd/yy")
   End If
   
   If iBarOnTop = 0 Then
      lLeft = frm.Left + frm.ActiveControl.Left
   Else
      lLeft = frm.ActiveControl.Left
   End If
   If iBarOnTop = 0 Then If lLeft > 6000 Then lLeft = lLeft - 1095
   If (lLeft + SysCalendar.Width) > (MdiSect.Width - 600) Then lLeft = lLeft - (SysCalendar.Width - frm.ActiveControl.Width + 300)
   
'      If iBarOnTop = 0 Then
'         SysCalendar.Move MdiSect.SideBar.Width + lLeft, frm.Top + _
'            (frm.ActiveControl.Top + frm.ActiveControl.Height + 1000 + iAdjust)
'      Else
'         SysCalendar.Move lLeft, frm.Top + (frm.ActiveControl.Top + _
'            MdiSect.TopBar.Height + frm.ActiveControl.Height + 1000 + iAdjust)
'      End If

   lTop = frm.Top + frm.ActiveControl.Top _
      + frm.ActiveControl.Height + iAdjust
   If frm.MDIChild Then
      If iBarOnTop = 0 Then
         lLeft = lLeft + MdiSect.SideBar.Width
         lTop = lTop + 850
      Else
         lTop = lTop + MdiSect.TopBar.Height + 850
      End If
   End If
   
   SysCalendar.Move lLeft, lTop

   bSysCalendar = True
   If IsDate(sDate) Then SysCalendar.Calendar1.Value = Format(sDate, "mm/dd/yy")
   
   'if parent form is modal, we must show this as modal too
   On Error Resume Next
   SysCalendar.Calendar1.Refresh
   DoEvents
   SysCalendar.Show
   If Err Then
      SysCalendar.Show vbModal
   End If
   'refresh it so that it doesn't blink out
   SysCalendar.Calendar1.Refresh
   'combo.Refresh
   
End Sub


Public Sub ShowCalendarEx(frm As Form, Optional iAdjust As Integer)

   'display date selection calendar
   
   Dim iAdder As Integer
   Dim lLeft As Long
   Dim lTop As Long
   Dim sDate As Date
   
   Dim combo As Control
   Set combo = frm.ActiveControl
   
   calEx = True
   
   If IsDate(frm.ActiveControl.Text) Then
      'frm.ActiveControl.AddItem frm.ActiveControl.Text
   Else
      'frm.ActiveControl.AddItem Format(Now, "mm/dd/yy")
      combo.Text = Format(Now, "mm/dd/yyyy")
   End If
   
   'set form to pass date back to
   Set SysCalendar.FromForm = frm

   'On Error Resume Next
   'See if there is a date in the combo
   If IsDate(frm.ActiveControl.Text) Then
      sDate = Format(frm.ActiveControl.Text, "mm/dd/yyyy")
   Else
      sDate = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
   If iBarOnTop = 0 Then
      lLeft = frm.Left + frm.ActiveControl.Left
   Else
      lLeft = frm.ActiveControl.Left
   End If
   If iBarOnTop = 0 Then If lLeft > 6000 Then lLeft = lLeft - 1095
   If (lLeft + SysCalendar.Width) > (MdiSect.Width - 600) Then lLeft = lLeft - (SysCalendar.Width - frm.ActiveControl.Width + 300)
   
'      If iBarOnTop = 0 Then
'         SysCalendar.Move MdiSect.SideBar.Width + lLeft, frm.Top + _
'            (frm.ActiveControl.Top + frm.ActiveControl.Height + 1000 + iAdjust)
'      Else
'         SysCalendar.Move lLeft, frm.Top + (frm.ActiveControl.Top + _
'            MdiSect.TopBar.Height + frm.ActiveControl.Height + 1000 + iAdjust)
'      End If

   lTop = frm.Top + frm.ActiveControl.Top _
      + frm.ActiveControl.Height + iAdjust
   If frm.MDIChild Then
      If iBarOnTop = 0 Then
         lLeft = lLeft + MdiSect.SideBar.Width
         lTop = lTop + 850
      Else
         lTop = lTop + MdiSect.TopBar.Height + 850
      End If
   End If
   
   SysCalendar.Move lLeft, lTop

   bSysCalendar = True
   If IsDate(sDate) Then SysCalendar.Calendar1.Value = Format(sDate, "mm/dd/yyyy")
   
   'if parent form is modal, we must show this as modal too
   On Error Resume Next
   SysCalendar.Calendar1.Refresh
   DoEvents
   SysCalendar.Show
   If Err Then
      SysCalendar.Show vbModal
   End If
   'refresh it so that it doesn't blink out
   SysCalendar.Calendar1.Refresh
   'combo.Refresh
   
End Sub


''Public Function ConvertHours(cTime As Currency) As String
''   'Converts real hours like 8.3 to time like 08:18
''   'Note systax - Pass a number, returns a string or variant
''   'sSomeChangedTime = ConvertHours(8.3)
''
''   Dim min As Integer
''   min = cTime * 60
''   ConvertHours = DateAdd("n", min, "1/1/1900")
''End Function
''

'gets password encryption (self documenting)

Public Function GetPassword(sPassword As String) As String
   Dim iList As Integer
   Dim K As Integer
   Dim n As Integer
   Dim sNewPw As String
   
   On Error Resume Next
   K = Len(Trim$(sPassword))
   n = 79
   For iList = 1 To K
      n = n + 1
      Mid$(sPassword, iList, 1) = Chr$(Asc(Mid$(sPassword, iList, 1)) - n)
   Next
   For iList = K To 1 Step -1
      sNewPw = sNewPw & Mid$(sPassword, iList, 1)
   Next
   GetPassword = sNewPw
   
End Function

'sets password encryption (self documenting)

Public Function SetPassword(sPassword As String) As String
   Dim K As Integer
   Dim n As Integer
   Dim iList As Integer
   Dim sNewPw As String
   
   On Error Resume Next
   K = Len(Trim$(sPassword))
   For iList = K To 1 Step -1
      sNewPw = sNewPw & Mid$(sPassword, iList, 1)
   Next
   n = 79
   For iList = 1 To K
      n = n + 1
      Mid$(sNewPw, iList, 1) = Chr$(Asc(Mid$(sNewPw, iList, 1)) + n)
   Next
   SetPassword = sNewPw
   
End Function

'Retrieves the list of most recent selections from
'the registry and files the ActiveBar.
'Might as well do some other load stuff too

Public Sub GetRecentList(sMdiSect As String)
   Dim iList As Integer
   On Error Resume Next
   MdiSect.tmePanel = Format(Time, "h:mm AM/PM")
   'User
   If Trim(cUR.CurrentUser) = "" Then cUR.CurrentUser = GetSetting("Esi2000", "system", "UserId", cUR.CurrentUser)
   'DSN for Crystal - Make one if required
   'GetCrystalDSN
   
   On Error GoTo modErr1
   'bNoCrystal = True
   For iList = 0 To 4
      sRecent(iList) = GetSetting("Esi2000", sMdiSect, "Recent" & Trim(str(iList)), sRecent(iList))
      If Len(Trim(sRecent(iList))) < 2 Then
         'Nothing there and hide it
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = False
      Else
         'There is an entry and show it
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Visible = True
         MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Caption = sRecent(iList)
      End If
   Next
   Exit Sub
   
modErr1:
   On Error GoTo 0
   
End Sub

'Set common MdiForm activation methods
'Give SQL Server a chance to check us in
'11/8/04 Insert KeyState

Public Sub ActivateSection(sCurrentSection As String)
   'initialize the section
   'returns True if successful
   
   Dim bList As Byte
   Dim bOpenForm As Byte
   Dim iInsState As Integer
   Dim iState As Integer
   Dim sState As String
   
   Dim iRed As Integer
   Dim iGreen As Integer
   Dim iBlue As Integer
   
   
   iBarOnTop = GetSetting("Esi2000", "Programs", "BarOnTop", iBarOnTop)
   ShowHideTopBar
'   iRed = GetSetting("Esi2000", "System", "SectionBackColorR", iRed)
'   iGreen = GetSetting("Esi2000", "System", "SectionBackColorG", iGreen)
'   iBlue = GetSetting("Esi2000", "System", "SectionBackColorB", iBlue)
'   If iRed + iGreen + iBlue = 0 Then
'      MdiSect.BackColor = vbApplicationWorkspace
'   Else
'      MdiSect.BackColor = RGB(iRed, iGreen, iBlue)
'   End If
   
   'MDISect.BackColor = &HFFFFE0 'SKY BLUE
   MdiSect.BackColor = GetBackgroundColor()
   
   On Error Resume Next
   'image is in BitMaps ES200n.bmp
   'ES200n.BackColor is custom RGB(212,208,200)
   'ES200n.Font SerpentineDBo 12/Bold/ital
   'ES200n.Font Color is RGB(0,0,128)
   For bList = 0 To 10
      MdiSect.cmdSect(bList).BackColor = ES_SystemBackcolor
   Next
   MdiSect.cmdSect(bList).BackColor = ES_SystemBackcolor
   
   MdiSect.Logo(0).ToolTipText = "Key Software LLC"
   MdiSect.Logo(1).ToolTipText = "Key Software LLC"
   If MdiSect.Logo(0).Width < 1944 Then MdiSect.Logo(0).Width = 1944
   MdiSect.SideBar.Width = MdiSect.Logo(0).Width + 20
   MdiSect.Logo(0).Left = 0
   MdiSect.Logo(0).Top = (MdiSect.SideBar.Height * 0.8)
   MdiSect.Logo(1).Top = 0
   MdiSect.Logo(1).Left = (MdiSect.TopBar.Width * 0.8)
   'MDISect.RightBar.BackColor = ES_SystemBackcolor
   MdiSect.LeftBar.BackColor = ES_SystemBackcolor
   MdiSect.SideBar.BackColor = ES_SystemBackcolor
   MdiSect.TopBar.BackColor = ES_SystemBackcolor
   'MDISect.BotPanel.BackColor = ES_SystemBackcolor
   MdiSect.ActiveBar1.BackColor = ES_SystemBackcolor
   MdiSect.lblBotPanel.BackColor = ES_SystemBackcolor
   'MDISect.SystemMsg.BackColor = ES_SystemBackcolor
   MdiSect.OvrPanel.BackColor = ES_SystemBackcolor
   MdiSect.tmePanel.BackColor = ES_SystemBackcolor
   
   MdiSect.OvrPanel.Enabled = False
   MdiSect.OvrPanel = ""
   MdiSect.lblBotPanel = "Initializing."
   MdiSect.lblBotPanel.Refresh
   
' Done in the MDI selection form
'MM
'   If Not OpenSqlServer(False) Then
'      End
'   End If
   
   'Make sure that SQL Server is going to connect or fail
'BBS remarked out for ADO/RDO Conversion
'    If clsadocon.StillConnecting Then
'      MDISect.Enabled = False
'      Sleep 500
'   End If
   MdiSect.Enabled = True
   bUserAction = True
   MdiSect.lblBotPanel = "Ready.."
   Sleep 500
   MdiSect.lblBotPanel.FontItalic = False
   MdiSect.lblBotPanel = MdiSect.Caption
   MdiSect.SetFocus
   '     If sDataBase <> "Esi2000Db" Then
   '         User.Group1 = 1
   '         User.Group2 = 1
   '         User.Group3 = 1
   '         User.Group4 = 1
   '         User.Group5 = 1
   '         User.Group6 = 1
   '    End If
   bAutoCaps = GetSetting("Esi2000", "mngr", "AutoCaps", bAutoCaps)
   '1/29/04
   bOpenLastForm = GetSetting("Esi2000", "System", "Reopenforms", bOpenLastForm)
   
   Err.Clear
   
   Dim bInsState As Boolean
   bInsState = False
   
   'iInsState = GetKeyState(vbKeyInsert)
   Dim bytKeys(255) As Byte
   'Get status of the 256 virtual keys
   'GetKeyboardState bytKeys(0)
   'bInsState = bytKeys(VK_INSERT)
   
   GetKeyboardState bytKeys(0)
   'Change a key
   bytKeys(VK_INSERT) = 1
   'Set the keyboard state
   SetKeyboardState bytKeys(0)

   bInsState = GetKeyState(vbKeyInsert)
   
   If bInsState Then
      bInsertOn = True
      MdiSect.OvrPanel = "INSERT"
      MdiSect.OvrPanel.ToolTipText = "Insert Text Is On (Click me)"
   Else
      bInsertOn = False
      MdiSect.OvrPanel = "OVER"
      MdiSect.OvrPanel.ToolTipText = "Overtype Text Is On (Click me) "
   End If
      
   '11/10/04 sets Insert On as default unless user
   'exists with Insert Off (Global for this Workstation)
'   sState = GetSetting("Esi2000", "mngr", "InsertState", sState)
'   If sState = "" Then iState = 1 Else iState = Abs(Val(sState))
'   If iInsState <> iState Then
'      If iState = 1 Then
'         ToggleInsertKey True
'      Else
'         ToggleInsertKey False
'      End If
'   Else
'      MDISect.OvrPanel.Enabled = True
'      If iState = 1 Then
'         bInsertOn = True
'         MDISect.OvrPanel = "INSERT"
'         MDISect.OvrPanel.ToolTipText = "Insert Text Is On (Click me) "
'      Else
'         bInsertOn = False
'         MDISect.OvrPanel = "OVER"
'         MDISect.OvrPanel.ToolTipText = "Overtype Text Is On (Click me)"
'      End If
'   End If
   MdiSect.Timer5.Enabled = True
End Sub

'Standard Unload as save for all sections
'MdiSect QueryUnload Event

Public Sub UnLoadSection(sMdiSect As String, sThisSection As String)
   Dim iList As Integer
   MdiSect.Timer1.Enabled = False
   MdiSect.Timer2.Enabled = False
   MdiSect.Timer3.Enabled = False
   MdiSect.Timer4.Enabled = False
   MdiSect.Timer5.Enabled = False
   
   'tell mom we are not here
   SaveSetting "Esi2000", "Sections", sMdiSect, 0
   'Save system KeyInsert
   SaveSetting "Esi2000", "mngr", "InsertState", Abs(bInsertOn)
   'save Recent list
   For iList = 0 To 4
      SaveSetting "Esi2000", sThisSection, "Recent" & Trim(str(iList)), Trim(MdiSect.ActiveBar1.Bands("mnuFile").Tools("FileRecent" & Trim(str(iList))).Caption)
   Next
   
   
End Sub

'Standardize then Resize of the MdiForm
'6/16/06 Blocked redundant bar

Public Sub ResizeSection()
   '   ShowHideTopBar
   MdiSect.tmePanel.Left = (MdiSect.BotPanel.Width - 850)
   MdiSect.OvrPanel.Left = (MdiSect.BotPanel.Width - 1650)
   
End Sub

'Standardize the Sub Main procedure

Public Sub MainLoad(sEsiSection As String)
   Dim lTime As Long
   Dim iSect As Integer
   Dim sString As String
   
   MouseCursor ccHourglass '13
   On Error Resume Next
   'Trap App to see if MOM is watching. Code for CurDir
   'to allow testing and programming
    bUserAction = RunningInIDE
   If Not bUserAction Then
      On Error Resume Next
      iSect = GetSetting("Esi2000", "sections", "EsiOpen", iSect)
      If iSect = 0 Then
         MdiSect.bUnloading = 1
         SysWarn.Show
         Sleep 3000
         Unload SysWarn
      End If
   
   'if running in IDE, run as SYSMGR
   Else
      cUR.CurrentUser = "ADMINISTRATOR"
      'cUR.CurrentUser = "SYSMGR"
      InitializePermissions Secure, 1
   End If
   sFilePath = GetSetting("Esi2000", "System", "FilePath", sFilePath)
   If sFilePath = "" Then sFilePath = App.Path & "\"
   'tell mom we are here
   If iSect = 1 Then SaveSetting "Esi2000", "Sections", sEsiSection, 1
   sInitials = Trim(GetSetting("Esi2000", "System", "UserInitials", sInitials))
   bUserAction = True
   MdiSect.lblBotPanel.FontItalic = True
   
   sHelpType = GetSetting("Esi2000", "System", "HelpType", sHelpType)
   If sHelpType = "" Then sHelpType = "chm"
   If sHelpType = "chm" And Dir("c:\Program Files\ES2000\ES2000.chm") <> "" Then
      sHelpType = "chm"
      MdiSect.Cdi.HelpFile = sFilePath & "ES2000.hlp"
      App.HelpFile = "c:\Program Files\ES2000\ES2000.chm"
      bSysHelp = 1
   Else 'MM GetFavorites "EsiProd"
'MM sProgName = "Production"

      If Dir(sFilePath & "ES2000.hlp") <> "" Then
         sHelpType = "hlp"
         MdiSect.Cdi.HelpFile = sFilePath & "ES2000.hlp"
         App.HelpFile = sFilePath & "ES2000.hlp"
         bSysHelp = 1
      Else
         bSysHelp = 0
      End If
   End If

   iSect = Val(ScreenResolution())
   MouseCursor ccDefault
   
End Sub


Public Sub FillStates(frm As Form)
   'Fills state codes

   Dim ADOSte As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT STATECODE,STATEDEFAULT FROM CsteTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOSte)
   If bSqlRows Then
      With ADOSte
         On Error Resume Next
         Do Until .EOF
            'for vendors..
            If frm.Name = "PurcPRe03a" Then
               AddComboStr frm.cmbPste.hwnd, "" & Trim(!STATECODE)
               If !STATEDEFAULT = 1 Then
                  frm.cmbSte = "" & Trim(!STATECODE)
               End If
            End If
            If frm.Name = "SaleSLe03a" Then
               AddComboStr frm.cmbStSte.hwnd, "" & Trim(!STATECODE)
               AddComboStr frm.cmbBtSte.hwnd, "" & Trim(!STATECODE)
               If !STATEDEFAULT = 1 Then
                  frm.cmbStSte = "" & Trim(!STATECODE)
                  frm.cmbBtSte = "" & Trim(!STATECODE)
               End If
            End If
            AddComboStr frm.cmbSte.hwnd, "" & Trim(!STATECODE)
            If !STATEDEFAULT = 1 Then
               frm.cmbSte = "" & Trim(!STATECODE)
            End If
            .MoveNext
         Loop
         ClearResultSet ADOSte
      End With
   End If
   Set ADOSte = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillstates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume modErr2
modErr2:
   If Left(CurrError.Description, 5) = "S0002" Then
      CurrError.Number = 0
      CurrError.Description = ""
   Else
      DoModuleErrors frm
   End If
   
End Sub

Public Function GetBackgroundColor() As Long
    GetBackgroundColor = vbApplicationWorkspace     ' default
    Dim rs As ADODB.Recordset
    Dim hexString As String
    Dim hex As Long
    On Error GoTo modErr1
    hexString = Trim(GetConfUserSetting(USERSETTING_BackgroundColorRGB))
    If Len(hexString) = 6 Then
             hexString = Mid(hexString, 5, 2) & Mid(hexString, 3, 2) & Mid(hexString, 1, 2)
             hex = Val("&H" & hexString & "&")
             GetBackgroundColor = hex
    End If
    
    Exit Function
   
modErr1:
End Function



'Remove spaces (32), dashes (45) and tabs (9) from indexed fields to
'avoid duplicate entries
'Optionally trim the length of the entry
'8/14/99 added optional ES_IGNOREDASHES to compress leaving dashes
'        note requires length if used
'11/29/05 Changed procedure to use Replace

Public Function Compress(TestNo As Variant, Optional iLength As Integer, Optional bIgnoreDashes As Byte) As String
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(TestNo)
   If Len(PartNo) > 0 Then
      NewPart = Replace(PartNo, Chr$(9), "")    'tab
      NewPart = Replace(NewPart, Chr$(10), "")  'lf
      NewPart = Replace(NewPart, Chr$(13), "")  'cr
      NewPart = Replace(NewPart, Chr$(32), "")  'space
      NewPart = Replace(NewPart, Chr$(39), "")  'single quote
      If bIgnoreDashes = 0 Then NewPart = Replace(NewPart, Chr$(45), "")   'dash
   End If
   Compress = NewPart
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   Compress = TestNo
   
End Function

'' Calls the windows API to get the windows directory and
'' ensures that a trailing dir separator is present
'' Returns: The windows directory
'Public Function GetWindowsDir()
'    Dim intZeroPos   As Integer
'    Dim gintMAX_SIZE As Integer
'    Dim strBuf       As String
'    gintMAX_SIZE = 255  'Maximum buffer size
'
'    strBuf = Space$(gintMAX_SIZE)
'    'Get the windows directory and then trim the buffer to the exact length
'    'returned and add a dir sep (backslash) if the API didn't return one
'    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
'        intZeroPos = InStr(strBuf, Chr$(0))
'        If intZeroPos > 0 Then strBuf = Left$(strBuf, intZeroPos - 1)
'        GetWindowsDir = strBuf
'    Else
'        GetWindowsDir = ""
'    End If
'
'End Function



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


Public Sub FormInitialize()
   Dim a As Long
   Dim iToolsCount As Integer
   On Error GoTo modErr1
   bUserAction = True
   'Bring some order to the user. Like it or not 10/2/01
   LockWindowUpdate MdiSect.hwnd
   'ES_SystemBackcolor = RGB(212, 208, 200)
   MdiSect.ActiveBar1.BackColor = ES_SystemBackcolor
   For iToolsCount = 0 To MdiSect.Controls.count - 1
      If TypeOf MdiSect.Controls(iToolsCount) Is Label Then
         MdiSect.Controls(iToolsCount).BackColor = ES_SystemBackcolor
         MdiSect.Controls(iToolsCount).ForeColor = Es_TextForeColor
      End If
   Next
   LockWindowUpdate 0
   MdiSect.tmePanel.Left = (MdiSect.BotPanel.Width - 850)
   MdiSect.OvrPanel.Left = (MdiSect.BotPanel.Width - 1650)
   
   'set size when not minimized or maximized
   Dim State As Long
   State = MdiSect.WindowState
   MdiSect.WindowState = crptNormal
   MdiSect.Left = 10
   MdiSect.Top = 10
   MdiSect.Width = Screen.Width - 100
   MdiSect.Height = Screen.Height - 100
   MdiSect.WindowState = State
   'MdiSect.Crw.WindowState = crptNormal
   Exit Sub
   
modErr1:
   On Error GoTo 0
End Sub


Public Sub ActivityDocument()
   'Activity Codes
   'Inventory activity (InvaTable)
   '   1 = Beginning Balance (first established)
   '  19 = Manual Adjustments
   '  30 = ABC Cycle Count (3/26/04)
   '  32 = Inter Company Transfer
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
   '  16 = Canceled PO Item & PO Item Receipt
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
   '         & "WHERE PARTREF='" & Trim(Str(vItems(i, 3))) & "' "
   '     clsAdoCon.ExecuteSQL sSql
   '     AverageCost LTrim(Str(vItems(i, 3)))
   
   'Add to Activity
   '    sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2," _
   '        & "INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT) " _
   '        & "VALUES(25,'" & vItems(i, 3) & "','PACKING SLIP'," _
   '        & "'" & vItems(i, 0) & Trim(vItems(i, 1)) & "'," _
   '        & "'" & format(es_sysdate, "mm/dd/yy") & "'," & Val(vItems(i, 2)) & "," _
   '        & Val(vItems(i, 4)) & ",'" & sCreditAcct & "','" & sDebitAcct & "')"
   '    clsAdoCon.ExecuteSQL sSql
   '----------
End Sub

'Format for name fields etc
'See constant ES_FIRSTWORD for First word only
'9/29/04 Added Caps Lock

Public Function StrCase(sTextStr As Variant, Optional bTextOption As Byte)
   Dim iKeyState As Integer
   Dim iStrLen As Integer
   Dim sNewStr As String
   iKeyState = GetKeyState(vbKeyCapital)
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
   End If
   
End Function


Public Sub GetSystemMessage()
'   Static sOldMessage As String
'   Dim b As Byte
'   Dim RdoMsg As ADODB.recordset
'
'   MDISect.Timer5.Enabled = False
'   On Error GoTo ModErr1
'   sSql = "Qry_GetSysMessage"
'   Set RdoMsg = clsadocon.OpenResultset(sSql, rdOpenForwardOnly)
'   If Not RdoMsg.BOF And Not RdoMsg.EOF Then
'      With RdoMsg
'         If sOldMessage <> "" & Trim(!ALERTMSG) Then
'            MDISect.SystemMsg = "" & Trim(!ALERTMSG)
'            sOldMessage = "" & Trim(!ALERTMSG)
'         End If
'         ClearResultSet RdoMsg
'      End With
'   End If
'   Set RdoMsg = Nothing
'   MDISect.Timer5.Enabled = True
'   Exit Sub
'
'ModErr1:
'   MDISect.Timer5.Enabled = True
'   On Error GoTo 0
'
End Sub


'Find the post date for a period
'txtDte = GetPostDate(Me, txtDte)

Public Function GetPostDate(frm As Form, sDate As String) As String
   Dim ADODte As ADODB.Recordset
   Dim b As Byte
   Dim iList As Integer
   
   On Error GoTo modErr1
   For iList = 1 To 13
      sSql = "SELECT FYYEAR,FYPERSTART" & Trim(str(iList)) & "," _
             & "FYPEREND" & Trim(str(iList)) & " FROM GlfyTable WHERE ('" _
             & sDate & "' BETWEEN FYPERSTART" & Trim(str(iList)) _
             & " AND FYPEREND" & Trim(str(iList)) & ") "
      bSqlRows = clsADOCon.GetDataSet(sSql, ADODte, ES_FORWARD)
      If bSqlRows Then
         GetPostDate = Format(ADODte.Fields(2), "mm/dd/yy")
         Exit For
      End If
   Next
   If GetPostDate = "" Then
      MsgBox "No Posting Date In The Period Selected.", _
         vbInformation, frm.Caption
   End If
   Set ADODte = Nothing
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
   Dim ADOFrm As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT PreRecord,PrePackSlip,PreInvoice,PrePurchaseOrder," _
          & "PreStateMent FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOFrm, ES_FORWARD)
   If bSqlRows Then
      With ADOFrm
         Select Case UCase$(Compress(sForm))
            Case "PACKSLIP"
               GetPrintedForm = .Fields(1)
            Case "INVOICE"
               GetPrintedForm = .Fields(2)
            Case "PURCHASEORDER"
               GetPrintedForm = .Fields(3)
            Case "STATEMENT"
               GetPrintedForm = .Fields(4)
            Case Else
               GetPrintedForm = 0
         End Select
      End With
   End If
   Set ADOFrm = Nothing
   
modErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
   
End Function

'Use local errors
'Execute direct for SQL Server...note stop on "'" (ANSI 39)
'See New ReplaceString
'11/21/06 Revised to use Replace

Public Function CheckComments(sComments As String) As String
   CheckComments = Replace(sComments, Chr$(39), Chr$(180))
   
End Function

Public Function GetTimeOut(sLastTime As String) As String
   GetTimeOut = "Last Access " & sLastTime & ", Timeout " _
                & Format(Time, "hh:mm AM/PM") & vbCrLf _
                & "Normal Database Connection Timeout." & vbCrLf _
                & "Reconnect To Service?"
   
End Function



'Need to watch what the enter

Public Function CheckValidColumn(sColumn As Variant) As Boolean
   Dim iPos As Integer
   Dim K As Integer
   Dim g As Byte
   Dim b As Byte
   Dim sRefCol As String
   
   On Error GoTo modErr1
   sRefCol = Trim(sColumn)
   iPos = Len(sRefCol)
   If iPos > 0 Then
      For K = 1 To iPos
         If Mid$(sRefCol, K, 1) <> Chr$(32) Then
            If Mid$(sRefCol, K, 1) < Chr$(46) Then
               b = 1
               g = Asc(Mid$(sRefCol, K, 1))
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
   Resume modErr2
modErr2:
   On Error Resume Next
   CheckValidColumn = False
   
   
End Function


''Allows selection of printers for individual reports
''Crystal requires this stuff
'Public Function GetPrinterPort(devPrinter As String, devDriver As String, devPort As String) As Byte
'    Dim SysPrinter As Printer
'        For Each SysPrinter In Printers
'            If Trim(SysPrinter.DeviceName) = devPrinter Then
'                devDriver = SysPrinter.DriverName
'                devPort = SysPrinter.Port
'                Exit For
'            End If
'        Next
'
'End Function
'
'New for VB6.0 starting 1/18/01
'   Note: DOES NOT work in Control Arrays
'   FormatControls Syntax in every form
'   Dim b As Byte
'   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())

Public Function AutoFormatControls(frm, TKeyPress() As EsiKeyBd, TGotFocus() _
                                   As EsiKeyBd, TKeyDown() As EsiKeyBd) As Byte
   ' //Need the following in case of a untrapped Control Array.
   ' Manual Code those from Module Procedures.
   Dim bByte As Byte
   Dim iRow As Integer
   Dim iList As Integer
   Dim b As Integer
   Dim c As Integer
   Dim n As Integer
   
   Dim ESI_txtKeyPress() As New EsiKeyBd
   Dim ESI_txtGotFocus() As New EsiKeyBd
   Dim ESI_txtKeyDown() As New EsiKeyBd
   
   iList = -1
   b = -1
   LockWindowUpdate frm.hwnd
   'Have to allow for arrays, etc-blast thru
   On Error Resume Next
   For iRow = 0 To frm.Controls.count - 1
   
'      If TypeOf frm.Controls(iRow) Is ComboBox Then
'         If frm.Controls(iRow).Style = 2 And frm.Controls(iRow).Tag = "9" Then   '2=dropdown list
'            'don't use Colin's stuff for dropdown lists with tag = 9
'            'Debug.Print "Don't autoformat " & frm.Controls(iRow).Name
'            iList = iList + 1
'            ReDim Preserve ESI_txtKeyPress(iList) As New EsiKeyBd
'            Set ESI_txtKeyPress(iList).esTxtKeyCheck = frm.Controls(iRow)
'         Else

         If True Then
            If True Then
            
'If frm.Controls(iRow).Name = "txtEnd" Then
'   Debug.Print "found"
'End If
            
            'Part of an Array or label (z1(n))?
            c = frm.Controls(iRow).Index
            If TypeOf frm.Controls(iRow) Is CommandButton Then
               If frm.Controls(iRow).Name = "ShowPrinters" Then
                  Set ESI_cmdShowPrint.esCmdClick = frm.Controls(iRow)
               End If
            End If
            If Err > 0 And (TypeOf frm.Controls(iRow) Is TextBox Or _
                            TypeOf frm.Controls(iRow) Is ComboBox Or TypeOf frm.Controls(iRow) Is MaskEdBox) Then
'If frm.Controls(iRow).Name = "txtCmt" Then
'   Debug.Print "found it"
'End If
               Err.Clear
               iList = iList + 1
               ReDim Preserve ESI_txtKeyPress(iList) As New EsiKeyBd
               If frm.Controls(iRow).Tag <> "9" Then
                  b = b + 1
                  ReDim Preserve ESI_txtGotFocus(b) As New EsiKeyBd
                  ReDim Preserve ESI_txtKeyDown(b) As New EsiKeyBd
               End If
               If TypeOf frm.Controls(iRow) Is MaskEdBox Then
                  Set ESI_txtGotFocus(b).esMskGotFocus = frm.Controls(iRow)
                  Set ESI_txtKeyDown(b).esMskKeyDown = frm.Controls(iRow)
                  Set ESI_txtKeyPress(iList).esMskKeyValue = frm.Controls(iRow)
               End If
               If TypeOf frm.Controls(iRow) Is TextBox Then
                  bByte = True
                  Select Case Val(frm.Controls(iRow).Tag)
                     Case 1
                        Set ESI_txtKeyPress(iList).esTxtKeyValue = frm.Controls(iRow)
                     Case 3
                        Set ESI_txtKeyPress(iList).esTxtKeyCase = frm.Controls(iRow)
                     Case 4
                        Set ESI_txtKeyPress(iList).esTxtKeyDate = frm.Controls(iRow)
                     Case 5
                        Set ESI_txtKeyPress(iList).esTxtKeyTime = frm.Controls(iRow)
                     Case 9
                        Set ESI_txtKeyPress(iList).esTxtKeyMemo = frm.Controls(iRow)
                        bByte = False
                     Case Else
                        Set ESI_txtKeyPress(iList).esTxtKeyCheck = frm.Controls(iRow)
                  End Select
                  If bByte Then
                     Set ESI_txtGotFocus(b).esTxtGotFocus = frm.Controls(iRow)
                     Set ESI_txtKeyDown(b).esTxtKeyDown = frm.Controls(iRow)
                  End If
               Else
                  If TypeOf frm.Controls(iRow) Is ComboBox Then
                     Set ESI_txtGotFocus(b).esCmbGotfocus = frm.Controls(iRow)
                     Select Case Val(frm.Controls(iRow).Tag)
                        Case 1
                           Set ESI_txtKeyPress(iList).esCmbKeyValue = frm.Controls(iRow)
                        Case 2
                           Set ESI_txtKeyPress(iList).esCmbKeyCheck = frm.Controls(iRow)
                        Case 3
                           Set ESI_txtKeyPress(iList).esCmbKeyCase = frm.Controls(iRow)
                        Case 4
                           Set ESI_txtKeyPress(iList).esCmbKeyDate = frm.Controls(iRow)
                        Case 8
                           Set ESI_txtKeyPress(iList).esCmbKeylock = frm.Controls(iRow)
                           frm.Controls(iRow).ForeColor = ES_BLUE
                        Case 9
                           Set ESI_txtKeyPress(iList).esCmbDropdownList = frm.Controls(iRow)
                           frm.Controls(iRow).ForeColor = ES_BLUE
                        Case Else
                           Set ESI_txtKeyPress(iList).esCmbKeyCase = frm.Controls(iRow)
                     End Select
                  End If
               End If
            End If
            Err.Clear
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
   LockWindowUpdate 0
   
End Function

Public Function GetLastActivity() As Long
   Dim ADOAct As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(INNUMBER) FROM InvaTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOAct, ES_FORWARD)
   If bSqlRows Then
      With ADOAct
         If Not IsNull(.Fields(0)) Then
            GetLastActivity = .Fields(0)
         Else
            GetLastActivity = 0
         End If
         ClearResultSet ADOAct
      End With
   End If
   Set ADOAct = Nothing
   Exit Function
modErr1:
   GetLastActivity = 0
   
End Function

Public Function IllegalCharacters(TestInputString As Variant) As Byte
   Dim iLen As Integer
   Dim K As Integer
   Dim sString As String
   
   ' check for illegal characters in part name: # $ % * * , / : ; @ '
   
   On Error GoTo modErr1
   sString = Trim$(TestInputString)
   iLen = Len(sString)
   IllegalCharacters = 0
   If iLen > 0 Then
      For K = 1 To iLen
         If Mid$(sString, K, 1) = Chr$(33) Or Mid$(sString, K, 1) = Chr$(34) _
                   Or Mid$(sString, K, 1) = Chr$(35) Or Mid$(sString, K, 1) = Chr$(36) _
                   Or Mid$(sString, K, 1) = Chr$(38) Or Mid$(sString, K, 1) = Chr$(42) _
                   Or Mid$(sString, K, 1) = Chr$(44) Or Mid$(sString, K, 1) = Chr$(47) _
                   Or Mid$(sString, K, 1) = Chr$(58) Or Mid$(sString, K, 1) = Chr$(59) _
                   Or Mid$(sString, K, 1) = Chr$(64) Or Mid$(sString, K, 1) = Chr$(37) _
                   Or Mid$(sString, K, 1) = Chr$(39) Or Mid$(sString, K, 1) = Chr$(146) _
                   Then
            IllegalCharacters = Asc(Mid$(sString, K, 1))
            Exit For
         End If
      Next
   End If
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   IllegalCharacters = 0
   
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
   Dim iPos As Integer
   Dim l As Long
   Dim dblTime As Double
   Dim sTime As String
   Static bNextLot As Variant         'Cycle lots

   On Error Resume Next
   ES_SYSDATE = GetServerDateTime()
   Randomize
   iPos = Int((99 * Rnd) + 1)
   If IsEmpty(bNextLot) Then
      bNextLot = Int((99 * Rnd) + 1)
   Else
      bNextLot = (bNextLot + 1) Mod 100
   End If
   dblTime = TimeValue(Format(ES_SYSDATE, "hh:nn:ss"))
   sTime = sTime & "-" & Format$(dblTime, ".000000") & Format$(bNextLot, "00")
   sTime = Right$(sTime, 6)
   dblTime = DateValue(Format(ES_SYSDATE, "mm/dd/yy"))
   sTime = Format$(dblTime, "00000") & "-" & sTime
   sTime = sTime & "-" & Format$(Trim$(str$(iPos)), "00")
   GetNextLotNumber = Trim(sTime)
   'If bNextLot > 99 Then bNextLot = 0
   'For l = 0 To 128000
   'Next 'Give some time without using sleep
   
End Function

Public Function CheckLotTracking() As Byte
   'See if the Lots are registered
   'Change later to see if Lot Tracking is active
   '3/12/02

   Dim ADOLots As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT LOTNUMBER FROM LohdTable WHERE LOTNUMBER=''"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOLots, ES_FORWARD)
   If bSqlRows Then ClearResultSet ADOLots
   CheckLotTracking = 1
   Set ADOLots = Nothing
   Exit Function
   
modErr1:
   On Error GoTo 0
   CheckLotTracking = 0
   
End Function

'3/26/02 FIFI or LIFO - Default is FIFO
'1 = FIFO 0 = LIFO

Public Function GetInventoryMethod() As Byte
   Dim ADOInm As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT COLOTSFIFO FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOInm, ES_FORWARD)
   If bSqlRows Then
      With ADOInm
         If Not IsNull(!COLOTSFIFO) Then
            GetInventoryMethod = !COLOTSFIFO
         Else
            GetInventoryMethod = 1
         End If
         ClearResultSet ADOInm
      End With
   End If
   Set ADOInm = Nothing
   Exit Function
   
modErr1:
   GetInventoryMethod = 1
   
End Function

Public Function UsingFifo() As Boolean
   'determine whether using FIFO inventory allocation
   Dim Ado As ADODB.Recordset
   sSql = "SELECT ISNULL(COLOTSFIFO, 1) as COLOTSFIFO FROM ComnTable WHERE COREF=1"
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      If Ado!COLOTSFIFO = 0 Then
         Exit Function
      End If
   End If
   UsingFifo = True
   Set Ado = Nothing
End Function

Public Function GetNextLotRecord(sCurrentLot As String) As Long
   ' Retrieves the next lot record for Lot Tracking
   Dim ADOLor As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(LOIRECORD) FROM LoitTable WHERE " _
          & "LOINUMBER='" & sCurrentLot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOLor, ES_FORWARD)
   If bSqlRows Then
      With ADOLor
         If Not IsNull(.Fields(0)) Then
            GetNextLotRecord = .Fields(0) + 1
         Else
            GetNextLotRecord = 2
         End If
         ClearResultSet ADOLor
      End With
   Else
      GetNextLotRecord = 2
   End If
   Set ADOLor = Nothing
   Exit Function
   
modErr1:
   GetNextLotRecord = 2
   
End Function


Public Function CheckLotStatus() As Byte
   Dim ADOLots As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT COLOTSACTIVE FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOLots, ES_FORWARD)
   If bSqlRows Then
      With ADOLots
         If Not IsNull(!COLOTSACTIVE) Then
            CheckLotStatus = !COLOTSACTIVE
         Else
            CheckLotStatus = 0
         End If
      End With
   Else
      CheckLotStatus = 0
   End If
   Set ADOLots = Nothing
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   CheckLotStatus = 0
   
End Function

Public Function GetRemainingLotQty(LotPart As String, _
                                   Optional UpdateHeader As Boolean) As Currency
   Dim ADOQty As ADODB.Recordset
   Dim iList As Integer
   Dim iRows As Integer
   Dim LotsAvailLot(900) As String
   Dim LotsAvailQty(900) As Currency

   On Error GoTo modErr1
   sSql = "SELECT LOTNUMBER,SUM(LOIQUANTITY) As LitemQty FROM " _
          & "LohdTable,LoitTable WHERE (LOTNUMBER=LOINUMBER AND " _
          & "LOIPARTREF='" & LotPart & "' AND LOTAVAILABLE=1) GROUP BY LOTNUMBER"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOQty, ES_FORWARD)
   If bSqlRows Then
      With ADOQty
         Do Until .EOF
            If Not IsNull(!LitemQty) Then
               If iRows < 900 Then
                  iRows = iRows + 1
                  LotsAvailLot(iRows) = "" & Trim(!lotNumber)
                  LotsAvailQty(iRows) = !LitemQty
               End If
               GetRemainingLotQty = (GetRemainingLotQty + !LitemQty)
            End If
            .MoveNext
         Loop
         ClearResultSet ADOQty
      End With
      
      '7/17/08 Release 55.  don't do the following -- anaylze the problem with the part quantity health report
      'and take appropriate action
'      If UpdateHeader = True Then
'         For iList = 1 To iRows
'            sSql = "UPDATE LohdTable SET LOTREMAININGQTY=" _
'                   & LotsAvailQty(iList) & " WHERE LOTNUMBER='" _
'                   & LotsAvailLot(iList) & "'"
'            ADOCon.Execute sSql, rdExecDirect
'         Next
'      End If
   End If
   Erase LotsAvailLot
   Erase LotsAvailQty
   Set ADOQty = Nothing
   Exit Function

modErr1:
   GetRemainingLotQty = 0

End Function


Public Function GetLotRemainingQty(LotPart As String) As Currency

   'same as GetRemainingLotQty, except gets remaining qty from sum(LOTREMAININGQTY)
   'rather than SUM(LOIQUANTITY) to reduce a problem at LUMICOR
   
   Dim ADOQty As ADODB.Recordset
   
   GetLotRemainingQty = 0
   sSql = "select isnull(sum(LOTREMAININGQTY),0)" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "where LOTPARTREF='" & LotPart & "' AND LOTAVAILABLE=1"
   If clsADOCon.GetDataSet(sSql, ADOQty, ES_FORWARD) Then
      GetLotRemainingQty = ADOQty.Fields(0)
   End If
   Set ADOQty = Nothing
   
End Function


Public Function GetLotTrailer() As String
   Static bCounter As Byte
   If bCounter = 0 Or bCounter > 89 Then bCounter = 64
   bCounter = bCounter + 1
   GetLotTrailer = Chr(bCounter)
   
End Function

'For Routing/Mo Times
'ES_TimeFormat = GetTimeFormat()

Public Function GetTimeFormat() As String
   Dim ADOFmt As ADODB.Recordset
   
   On Error GoTo modErr1
   sSql = "SELECT TimeFormat FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOFmt, ES_FORWARD)
   If bSqlRows Then
      With ADOFmt
         If Not IsNull(.Fields(0)) Then
            GetTimeFormat = "" & Trim(.Fields(0))
         Else
            GetTimeFormat = ES_QuantityDataFormat
         End If
         ClearResultSet ADOFmt
      End With
   Else
      GetTimeFormat = ES_QuantityDataFormat
   End If
   Set ADOFmt = Nothing
   Exit Function
   
modErr1:
   GetTimeFormat = ES_QuantityDataFormat
   On Error GoTo 0
   
End Function

Public Function GetPartQuantityOnHand(strPartRef As String) As Currency

   Dim rdo As ADODB.Recordset
   sSql = "SELECT ISNULL(PAQOH,0) AS PAQOH from PartTable WHERE PartRef = '" & strPartRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   
   If bSqlRows Then
      GetPartQuantityOnHand = Format(rdo!PAQOH, ES_QuantityDataFormat)
   Else
      GetPartQuantityOnHand = Format(0, ES_QuantityDataFormat)
   End If
   Set rdo = Nothing

End Function

'set the form as on top

Public Sub AlwaysOnTop(hwnd As Long, bTop As Boolean)
   If Not bTop Then
      SetWindowPos hwnd, Hwnd_NOTOPMOST, 0, 0, 0, 0, Swp_NOMOVE + Swp_NOSIZE
   Else
      SetWindowPos hwnd, hWnd_TopMost, 0, 0, 0, 0, Swp_NOMOVE + Swp_NOSIZE
   End If
   
End Sub

'New 7/11/03
'Current coding principles:
'MdiSect.Crw.ReportFileName = sReportPath & "invlt01.rpt"
'New coding principals ** sCustomReport is a Project varible **
'sCustomReport = GetCustomReport("invlt01")
'MdiSect.Crw.ReportFileName = sReportPath & sCustomReport
'
'Note that the Group doesn't matter.  Trying to keep in mind changes to
'the caption too.
'12/26/05 Revised (Added Replace)

Public Function GetCustomReport(StdReport As String) As String
   Dim ADOCst As ADODB.Recordset
   Dim sNewReport As String
   
   'Strip the extension
   StdReport = Replace(StdReport, ".rpt", "")
   sNewReport = Trim(LCase$(StdReport))
   
   'get module name.
   Dim module As String
   If InTestMode Then
      module = Mid(MdiSect.Caption, Len("TEST MODE ") + 1, 4)
   Else
      module = Left(MdiSect.Caption, 4)
   End If
   
   On Error GoTo modErr1
   
   'revised 4/10/2017 to only require one custom report definition for ALL modules rather
   'than one definition per module
'   sSql = "SELECT REPORT_INDEX,REPORT_REF,REPORT_SECTION,REPORT_CUSTOMREPORT " _
'          & "FROM CustomReports WHERE REPORT_SECTION LIKE '" & module & "%' " _
'          & "AND REPORT_REF='" & Compress(sNewReport) & "'"
   sSql = "SELECT REPORT_CUSTOMREPORT " _
          & "FROM CustomReports WHERE REPORT_REF='" & Compress(sNewReport) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, ADOCst, ES_FORWARD)
   If bSqlRows Then sNewReport = "" & Trim(ADOCst!REPORT_CUSTOMREPORT)
   If Trim(sNewReport) = "" Then sNewReport = LCase$(StdReport)
   GetCustomReport = sNewReport & ".rpt"
   Set ADOCst = Nothing
   Exit Function
   
modErr1:
   'it failed
   GetCustomReport = LCase$(StdReport) & ".rpt"
   
End Function

'2/6/04 If Trim(txtEml) <> "" Then SendEMail Trim(txtEml)

Public Sub SendEMail(SendTo As String)
   ShellExecute MdiSect.ActiveForm.hwnd, "open", "mailto:" & SendTo, _
      vbNullString, vbNullString, SW_SHOW
   '//ShellExecute(
   'hwnd>> handle to parent window
   'lpOperation>> pointer to string that specifies operation to perform
   'lpFile>> pointer to filename or folder name string
   'lpParameters>> pointer to string that specifies executable-file parameters
   'lpDirectory>> pointer to string that specifies default directory
   ' nShowCmd>>whether file is shown when opened
   ')
   '    MDISect.MAPISession1.SignOn
   '
   '    MDISect.MAPIMessages1.SessionID = MDISect.MAPISession1.SessionID
   '
   '    MDISect.MAPIMessages1.Compose
   '    MDISect.MAPIMessages1.RecipAddress = SendTo
   '    MDISect.MAPIMessages1.MsgSubject = "Here is my file"
   '    MDISect.MAPIMessages1.MsgNoteText = "Below you will find the file."
   '
   '    'Add the Attachment at the end of the message
   '    MDISect.MAPIMessages1.AttachmentPosition = 0
   '
   '    'Set the type to a data file
   '    MDISect.MAPIMessages1.AttachmentType = mapData
   '
   '    'Give it a name
   '    MDISect.MAPIMessages1.AttachmentName = "Invoice"
   '
   '    'Specify what file to send
   '    MDISect.MAPIMessages1.AttachmentPathName = "c:\download\prdpr02.rpt"
   '
   '    MDISect.MAPIMessages1.sEnd True
   '
   '    MDISect.MAPISession1.SignOff
End Sub

'2/9/04 If Trim(txtEml) <> "" Then OpenWebPage Trim(txtWeb)
'SendTo = request and doesn't have to be a web address
'sOperation to be "open" or "print"

Public Sub OpenWebPage(SendTo As String, Optional sOperation As String)
   Dim lTask As Long
   If sOperation = "" Then sOperation = "open"
   On Error Resume Next
   lTask = ShellExecute(MdiSect.hwnd, sOperation, SendTo, _
           vbNullString, vbNullString, SW_SHOW)
   If lTask < 32 Then MsgBox SendTo & " doesn't exist or couldn't be opened." & vbCrLf _
      & "If the file is a Log, then there may be no data to report.", _
      vbInformation, "ES/2000 ERP"
   
   '//ShellExecute(
   'hwnd>> handle to parent window
   'lpOperation>> pointer to string that specifies operation to perform
   'lpFile>> pointer to filename or folder name string
   'lpParameters>> pointer to string that specifies executable-file parameters
   'lpDirectory>> pointer to string that specifies default directory
   ' nShowCmd>>whether file is shown when opened
   'sOperation
   '   "Open"   The function opens the file specified by the lpFile parameter.
   '   The file can be an executable file or a document file. It can also be
   '   a folder.
   '
   '    "Print"  The function prints the file specified by lpFile. The file
   '    should be a document file. If the file is an executable file,
   '    the function opens the file, as if "open" had been specified.
   '
   '   "Explore"   The function explores the folder specified by lpFile.
   '   "Play"  For methods supporting a play function, such as sound files.
   '   "Properties"    For displaying the Properties page for files
   '   0&  lpOperation can be also be NULL (0& if declared As Any, or vbNullString if declared As String). In these cases the call performs the default verb action on the file specified, which is usually Open.  The default action can be seen by viewing specific extension in Explorer's Tool / Folder Options / File Types.
End Sub

'Syntax OpenHelpContext "4101" where the Context Id (page) is hs4101
'Note temporary URL
'8/1/04 New
'7/1/05 Revised to remove reference to WebHelp
'7/25/05 Cleared to show only .chm or .hlp Help
'        Added Registry setting to recall topic (FormLoad)

Public Sub OpenHelpContext(vContextID As Variant, Optional FromSect As Boolean)
   On Error GoTo ModErr
   'Note: Known as Topic Id to some Help producers
   If vContextID = 0 Then vContextID = 999
   If sHelpType = "chm" Then
      If FromSect Then
         'necessary to make the help show on an MdiForm
         '(bug in ActiveBar)
         SysOpen.HelpContextID = vContextID
         SysOpen.Left = -SysOpen.Width
         SysOpen.Show
         SendKeys "{F1}"
      Else
         'All others
         MdiSect.ActiveForm.HelpContextID = vContextID
         SendKeys "{F1}" 'Required if pressing cmdHlp
         'Next time user will not receive the default on Form.KeyDown
         SaveSetting "Esi2000", "Help", MdiSect.ActiveForm.Caption, vContextID
      End If
   Else
      If bSysHelp = 1 Then
         '.hlp does not require the SendKeys
         With MdiSect.Cdi
            .HelpContext = vContextID
            .HelpCommand = cdlHelpContext
            .ShowHelp
         End With
      End If
   End If
   Exit Sub
   
ModErr:
   If sHelpType = "chm" Then sHelpType = "hlp"
   MsgBox sSysCaption & " Help Is Not Installed Or Not In The Correct Folder.", _
      vbInformation, sSysCaption
   
End Sub

'10/2/04 Checks change DO NOT USE WITH index arrays

Public Function AutoFormatChange(frm, TChange() As EsiKeyBd) As Byte
   Dim iRow As Integer
   Dim iList As Integer
   Dim b As Integer
   Dim c As Integer
   Dim ESI_txtChange() As New EsiKeyBd
   
   For iRow = 0 To frm.Controls.count - 1
      On Error Resume Next
      c = frm.Controls(iRow).Index
      If Err > 0 And (TypeOf frm.Controls(iRow) Is TextBox Or _
                      TypeOf frm.Controls(iRow) Is ComboBox) Then
         b = b + 1
         ReDim Preserve ESI_txtChange(b) As New EsiKeyBd
         If (TypeOf frm.Controls(iRow) Is TextBox) Then
            Set ESI_txtChange(b).esTxtChange = frm.Controls(iRow)
         Else
            Set ESI_txtChange(b).esCmbChange = frm.Controls(iRow)
         End If
      End If
   Next
   
   TChange = ESI_txtChange()
   Erase ESI_txtChange
   
End Function

'11/13/04 New
'3/20/06 Set ListIndex
'LoadComboBox cmbPrt, -1    //Optional for Columns other than 0 (1st column)
'(rdoColumns are zero based)

Public Sub LoadComboBox(Cntrl As Control, Optional ColumnNumber As Integer, Optional SelectFirst As Boolean = True)
   'fill a combo box with the results of a query in sSql
   'For historic (stupid) reasons, the column is -1 based.
   'if you want the first column, pass -1
   'if no column is specified, the second column is used
   ' to not select any column pass in -2  CAUSES AN ERROR.  use -1 with SelectFirst = False
   
   Dim ComboLoad As ADODB.Recordset
   Cntrl.Clear
'Debug.Print "LoadComboBox " & Cntrl.Name & " Clear count = " & Cntrl.ListCount
   If sSql = "" Then Exit Sub
   ColumnNumber = ColumnNumber + 1
   Set ComboLoad = clsADOCon.GetRecordSet(sSql, ES_STATIC)
'   Set ComboLoad = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not ComboLoad.BOF And Not ComboLoad.EOF Then
      With ComboLoad
         Do Until .EOF
            AddComboStr Cntrl.hwnd, "" & Trim(.Fields(ColumnNumber))
            .MoveNext
         Loop
         ClearResultSet ComboLoad
      End With
   End If
   If Cntrl.ListCount <> 0 Then
      bSqlRows = 1
      If SelectFirst Then
         Cntrl.ListIndex = 0
      End If
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set ComboLoad = Nothing
   
End Sub

Public Sub LoadComboBoxAndSelect(cbo As ComboBox, Optional SelectString As String)

   ' load the combobox with the contents of the first column returned from ssql
   ' select the first entry >= the string specified
   ' if no string is specified, select the first entry
   
      
   Dim Ado As ADODB.Recordset
   cbo.Clear
   
   'text is read-only for dropdown list
   If cbo.Style <> 2 Then
      cbo.Text = ""
   End If
   If sSql = "" Then Exit Sub
   Set Ado = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
 
'   Set rdo = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not Ado.BOF And Not Ado.EOF Then
      bSqlRows = True
      With Ado
         Do Until .EOF
            If cbo.ListCount = 32766 Then
               AddComboStr cbo.hwnd, "MORE THAN 32767 ROWS"
               'cbo.ListIndex = cbo.ListCount - 1
               cbo.Text = "MORE THAN 32767 ROWS"
               Exit Do
            End If
            AddComboStr cbo.hwnd, "" & Trim(.Fields(0))
            If cbo.Text = "" Then
               If Trim(.Fields(0)) >= SelectString Then
                  If .Fields(0) <> "<ALL>" Then
                     'cbo.Text = Trim(.Fields(0))
                     cbo.ListIndex = cbo.ListCount - 1
                  End If
               End If
            End If
            .MoveNext
         Loop
         ClearResultSet Ado
      End With
   Else
      bSqlRows = False
   End If
   Set Ado = Nothing
   
   'If cbo.ListIndex = -1 And cbo.ListCount > 0 Then
   '   cbo.ListIndex = 0      'CAUSES COLLAPSE
   'End If
   
   
End Sub



'11/29/04 New
'3/20/06 Set ListIndex
'LoadNumComboBox cmbPon, "000000", 0    //Optional for Columns other than 1
'(rdoColumns are zero based)

Public Sub LoadNumComboBox(Cntrl As Control, ColumnFormat As String, _
                           Optional ColumnNumber As Integer)
   Dim ComboLoad As ADODB.Recordset
   Dim iRows As Integer
   If sSql = "" Then Exit Sub
   Set ComboLoad = clsADOCon.GetRecordSet(sSql, ES_STATIC)
'   Set ComboLoad = clsadocon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   If Not ComboLoad.BOF And Not ComboLoad.EOF Then
      With ComboLoad
         Do Until .EOF
            iRows = iRows + 1
            If iRows > 500 Then Exit Do
            AddComboStr Cntrl.hwnd, Format(.Fields(ColumnNumber), ColumnFormat)
            .MoveNext
         Loop
         ClearResultSet ComboLoad
      End With
   End If
   If Cntrl.ListCount > 0 Then
      bSqlRows = 1
      Cntrl.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   Cntrl.ToolTipText = "Contains Up To 500 Most Recent Entries"
   sSql = ""
   Set ComboLoad = Nothing
   
End Sub




'3/15/05 Allows for Sat and Sun in schedules

Public Function GetScheduledDate(SchedDate As Variant, SchedDays As Integer) As Variant
   Dim dNewDate As Date
   On Error Resume Next
   dNewDate = Format(SchedDate, "mm/dd/yy hh:mm")
   'Back Schedule
   If SchedDays < 0 Then
      SchedDays = Abs(SchedDays)
      If Format(dNewDate - SchedDays, "ddd") = "Sat" Then
         SchedDays = SchedDays + 1
      ElseIf Format(dNewDate - SchedDays, "ddd") = "Sun" Then
         SchedDays = SchedDays + 2
      End If
      dNewDate = Format(dNewDate - SchedDays, "mm/dd/yy hh:mm")
   Else
      'Forward Schedule
      If Format(dNewDate + SchedDays, "ddd") = "Sat" Then
         SchedDays = SchedDays + 2
      ElseIf Format(dNewDate + SchedDays, "ddd") = "Sun" Then
         SchedDays = SchedDays + 1
      End If
      dNewDate = Format(dNewDate + SchedDays, "mm/dd/yy hh:mm")
   End If
   GetScheduledDate = Format(dNewDate, "mm/dd/yy hh:mm")
   
End Function


Public Function GetPartAccounts(PartNumber As String, _
   DebitAccount As String, CreditAccount As String) As Byte
   'bByte = GetPartAccounts(Compress(txtPrt), sCreditAcct, sDebitAcct)
   
   Dim ADOAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   Dim compressedPart As String
   
   compressedPart = Compress(PartNumber)
   On Error GoTo modErr1
   
   DebitAccount = ""
   CreditAccount = ""
   
   sSql = "select dbo.fnGetPartCgsAccount ( '" & compressedPart & "' )"
   If clsADOCon.GetDataSet(sSql, ADOAct, ES_FORWARD) Then
      DebitAccount = "" & ADOAct.Fields(0)
   End If
   
   sSql = "select dbo.fnGetPartInvAccount ( '" & compressedPart & "' )"
   If clsADOCon.GetDataSet(sSql, ADOAct, ES_FORWARD) Then
      CreditAccount = "" & ADOAct.Fields(0)
   End If
   
   Set ADOAct = Nothing
   Exit Function
   
modErr1:
sProcName = "GetPartAccounts"
CurrError.Number = Err.Number
CurrError.Description = Err.Description
DoModuleErrors MdiSect.ActiveForm

End Function

Public Function ValidPartNumber(PartNumber As String) As Boolean
   Dim ADOAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   Dim compressedPart As String
   
   compressedPart = Compress(PartNumber)
   On Error GoTo modErr1
   
   
   sSql = "select * FROM PartTable WHERE PARTREF = '" & compressedPart & "' AND PAINACTIVE = 0 AND PAOBSOLETE = 0"
   If clsADOCon.GetDataSet(sSql, ADOAct, ES_FORWARD) Then
      ValidPartNumber = True
   Else
      ValidPartNumber = False
   End If
   
   Set ADOAct = Nothing
   Exit Function
   
modErr1:
sProcName = "ValidPartNumber"
CurrError.Number = Err.Number
CurrError.Description = Err.Description
DoModuleErrors MdiSect.ActiveForm

End Function

''Note: Skips over KeySets and Dynamic Cursors
'Public Sub ClearResultSet(RdoDataSet As ADODB.recordset)
'    If Not RdoDataSet.Updatable Then
'        Do While RdoDataSet.MoreResults
'        Loop
'        RdoDataSet.Cancel
'    End If
'
'End Sub
'
'11/18/05
'Called after each Insert
'UpdateWipColumns lSysCount
'where lSysCount is >= the first INNUMBER in the transaction

Public Sub UpdateWipColumns(InventoryRows As Long)
   Dim AdoInv As ADODB.Recordset

   Dim sWiplab As String
   Dim sWipMat As String
   Dim sWipExp As String
   Dim sWipOhd As String

   On Error GoTo modErr1
   sSql = "SELECT " _
          & "WIPLABACCT" _
          & ",WIPMATACCT" _
          & ",WIPEXPACCT" _
          & ",WIPOHDACCT " _
          & "FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoInv, ES_FORWARD)
   If bSqlRows Then
      With AdoInv
         sWiplab = "" & Trim(!WIPLABACCT)
         sWipMat = "" & Trim(!WIPMATACCT)
         sWipExp = "" & Trim(!WIPEXPACCT)
         sWipOhd = "" & Trim(!WIPOHDACCT)

         .Cancel
      End With
      ClearResultSet AdoInv
   End If
   sSql = "UPDATE InvaTable SET " _
          & "INCRLABACCT='" & sWiplab & "'," _
          & "INCRMATACCT='" & sWipMat & "'," _
          & "INCREXPACCT='" & sWipExp & "'," _
          & "INCROHDACCT='" & sWipOhd & "' " _
          & "WHERE INNUMBER>=" & InventoryRows & " "
   clsADOCon.ExecuteSql sSql
   Set AdoInv = Nothing
   Exit Sub

modErr1:
   Err.Clear

End Sub

Public Sub ShowHideTopBar()
   If iBarOnTop = 1 Then
      MdiSect.SideBar.Visible = False
      MdiSect.LeftBar.Visible = True
      MdiSect.TopBar.Visible = True
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Side"
   Else
      MdiSect.SideBar.Visible = True
      MdiSect.LeftBar.Visible = False
      MdiSect.TopBar.Visible = False
      MdiSect.ActiveBar1.Bands("Options").Tools("FavorBar").Caption = "Bar On Top"
   End If
   
   
End Sub
   
'Public Function Debugging() As Boolean
'   If InStr(1, Command, "/debug", vbTextCompare) Then
'      Debugging = True
'   Else
'      Debugging = False
'   End If
'End Function
'
'Public Function RunningBeta() As Boolean
'   'determine whether running beta featurs
'   'used initially to replace ES_CUSTOM = "PROPLA" with RunningBeta
'   If InStr(1, Command, "/beta", vbTextCompare) Then
'      RunningBeta = True
'   Else
'      RunningBeta = False
'   End If
'End Function
'
'For numbers Not used yet. Use Str(lNumber)

Public Sub AddComboNum(lhWnd As Long, lNumber As Long)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, _
      ByVal lNumber
   
End Sub

Public Function ShutdownTest() As Boolean
   'return = true if a shutdown has been authorized by the user
   'otherwise, shutdown is automatic if it occurs

   Static iTimer As Integer
   Static sLast As String
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim CloseApp As Long
   Dim sMsg As String
   Dim CurSection As String

   If bUserAction Then
      iTimer = 0
      bUserAction = False
      sLast = Format$(Time, "hh:mm AM/PM")
   Else
      iTimer = iTimer + 1
   End If
   If iTimer = 60 Then '57
      If Not bUserAction Then
         sMsg = GetTimeOut(sLast)
         On Error Resume Next
         bUserAction = True
         bByte = InStr(LTrim$(MdiSect.Caption), "-")
         CurSection = " " & Left$(MdiSect.Caption, bByte - 2)
         CloseForms
         'clsadocon.Close
         If DateDiff("n", CDate("5:00 PM"), CDate(Format(Now, "HH:MM:SS AM/PM"))) > 0 Then
         'If tmePanel > "4:50 PM" Then
            SaveSetting "Esi2000", "System", "CloseSection", App.Title
            CloseApp = FindWindow(vbNullString, "ESI CloseSections")
            If CloseApp = 0 Then
               If Dir(sFilePath & "EsiExit.exe") <> "" Then _
                      Shell sFilePath & "EsiExit.exe", vbNormalFocus
            Else
               AppActivate "ESI CloseSections", True
               SendKeys "% x", True
            End If
         End If
         bResponse = MsgBox(sMsg, ES_YESQUESTION, sSysCaption & CurSection)
         If bResponse = vbYes Then
            iTimer = 0
            bUserAction = True
            OpenDBServer True
         Else
            'Unload Me
            ShutdownTest = True
         End If
      End If
   End If
End Function

Public Sub FillMoPartCombo(cboPart As ComboBox, cboRun As ComboBox, _
   whereClause As String, Optional AllowNone As Boolean)
'Public Sub FillRuns(frm As Form, sSearchString As String, Optional sComboName As String)
   'fill MO part number combo box
   'use in conjunction with FillMoRunCombo
   'place the following in cboPart click event:
   '   FillMoPartInfo cboMoPart, lblPartDescription, lblPartType (optional)
   '   FillMoRunCombo cboMoPart, cboMoRun, whereClause

   Dim Ado As ADODB.Recordset
   On Error GoTo modErr1
   cboPart.Clear
   cboRun.Clear
   
'   If Not lblDescription Is Nothing Then
'      lblDescription.Caption = ""
'   End If

   If Not IsNull(AllowNone) Then
      If AllowNone Then
         cboPart.AddItem "<NONE>"
      End If
   End If
   
   sSql = "select distinct rtrim(PARTNUM) as PARTNUM from RunsTable" & vbCrLf _
      & "join PartTable on RUNREF = PARTREF" & vbCrLf _
      & whereClause & vbCrLf _
      & "order by rtrim(PARTNUM)"
   bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD)
   If bSqlRows Then
      With Ado
         Do Until .EOF
            cboPart.AddItem !PartNum
            .MoveNext
         Loop
      End With
   End If
   Set Ado = Nothing
   
   'set to first entry in list
   If cboPart.ListCount > 0 Then
      cboPart.ListIndex = 0
   End If
   Exit Sub
   
modErr1:
   sProcName = "FillMoPartCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub FillMoRunCombo(cboPart As ComboBox, cboRun As ComboBox, whereClause As String)
   'fill MO run number combo box with runs for current MO part
   
   'use in conjunction with FillMoPartCombo
   Dim Ado As ADODB.Recordset
   On Error GoTo modErr1
   cboRun.Clear
   sSql = "select RUNNO from RunsTable" & vbCrLf _
      & "join PartTable on RUNREF = PARTREF" & vbCrLf _
      & "and RUNREF = '" & Compress(cboPart.Text) & "'" & vbCrLf _
      & whereClause & vbCrLf _
      & "order by RUNNO"
   bSqlRows = clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD)
   If bSqlRows Then
      With Ado
         Do Until .EOF
            cboRun.AddItem !Runno
            .MoveNext
         Loop
      End With
   End If
   Set Ado = Nothing
   
   ' select first item
   If cboRun.ListCount > 0 Then
      cboRun.ListIndex = 0
      If GetPreferenceValue("AutoSelectLastRun") = "1" Then cboRun = cboRun.List(cboRun.ListCount - 1)
   End If
   
   Exit Sub
   
modErr1:
   sProcName = "FillMoRunCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Public Sub FillMoPartInfo(cboPart As ComboBox, lblDescription As Label, Optional lblType As Label)
   'update labels to display info for newly selected part in combobox
   
   If Not lblDescription Is Nothing Then
      lblDescription.Caption = ""
   End If
   
   If Not lblType Is Nothing Then
      lblType.Caption = ""
   End If
   
   Dim Ado As ADODB.Recordset
   sSql = "select PADESC, PALEVEL from PartTable" & vbCrLf _
      & "where PARTREF = '" & Compress(cboPart) & "'"
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      With Ado
         If Not lblDescription Is Nothing Then
            lblDescription.Caption = !PADESC
         End If
         
         If Not lblType Is Nothing Then
            lblType.Caption = !PALEVEL
         End If
      End With
   End If
   Set Ado = Nothing
End Sub


Public Sub FillRunInfo(cboPart As ComboBox, cboRun As ComboBox, lblStatus As Label, lblQty As Label)
   'update labels to display info for newly selected run in combobox
   
   If Not lblStatus Is Nothing Then
      lblStatus.Caption = ""
   End If
   
   If Not lblQty Is Nothing Then
      lblQty.Caption = ""
   End If
   
   Dim Ado As ADODB.Recordset
   sSql = "select RUNSTATUS, RUNQTY from RunsTable" & vbCrLf _
      & "where RUNREF = '" & Compress(cboPart) & "' AND RUNNO = " & cboRun
   If clsADOCon.GetDataSet(sSql, Ado, ES_FORWARD) Then
      With Ado
         If Not lblStatus Is Nothing Then
            lblStatus.Caption = !RUNSTATUS
         End If
         
         If Not lblQty Is Nothing Then
            lblQty.Caption = Format(!RUNQTY, "0.000")
         End If
      End With
   End If
   Set Ado = Nothing
   
End Sub

Public Function SavePictureToDB(sFileName As String, iRecID As Integer)

Dim bytData() As Byte

Dim intBlocks As Integer
Dim intBlocksLo As Integer
Dim lngImgLen As Long
Dim lngTxtLen As Long
Dim intCnt As Integer

Dim DataFile As Integer
Dim Fl As Long
Dim i As Integer
Dim sPicFile As String
Dim ADOPic As ADODB.Recordset
Const lngBlockSize As Integer = 15000

    
    DataFile = FreeFile
    sPicFile = sFileName
    On Error GoTo modErr1
    Open sPicFile For Binary Access Read As DataFile
    Fl = LOF(DataFile)
    
    sSql = "SELECT * FROM BitImage WHERE ImageRecord = " & CStr(iRecID)
    bSqlRows = clsADOCon.GetDataSet(sSql, ADOPic, ES_KEYSET)
   
    If (bSqlRows = False) Then
        ADOPic.AddNew
     'Else

    End If
    ' Update the logo image ID
    ADOPic!ImageRecord = iRecID
    
    intBlocks = Fl \ lngBlockSize
    intBlocksLo = Fl Mod lngBlockSize
    ADOPic!ImageStored.AppendChunk Null
    ReDim bytData(intBlocksLo)
    Get DataFile, , bytData()
    
    ADOPic!ImageStored.AppendChunk bytData()
    ReDim bytData(lngBlockSize)
    For i = 1 To intBlocks
       Get DataFile, , bytData()
       ADOPic!ImageStored.AppendChunk bytData()
    Next
    Close DataFile
    ADOPic.Update
    On Error Resume Next
    ADOPic.Close
    
    ' If every thing went well
    SavePictureToDB = True
    Set ADOPic = Nothing
    Exit Function
    
modErr1:
    sProcName = "SavePicture"
    CurrError.Number = Err.Number
    CurrError.Description = Err.Description
    DoModuleErrors MdiSect.ActiveForm
   
End Function


Public Function ReadImageFromDB(sFileName As String, iRecID As Integer) As Boolean

Dim bytData() As Byte
Dim strData As String
Dim intBlocks As Integer
Dim intBlocksLo As Integer
Dim lngImgLen As Long
Dim lngTxtLen As Long
Dim intCnt As Integer
Dim DataFile As Integer
Const lngBlockSize As Integer = 15000

Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim AdoParameter2 As ADODB.Parameter

Dim ADOPic As ADODB.Recordset

Set AdoQry = New ADODB.Command
AdoQry.CommandText = "{? = CALL Qry_ImageData(?) }"

Set AdoParameter1 = New ADODB.Parameter
AdoParameter1.Type = adInteger
AdoParameter1.Direction = adParamReturnValue

Set AdoParameter2 = New ADODB.Parameter
AdoParameter2.Direction = adParamInput
AdoParameter2.Type = adInteger
AdoParameter2.Value = iRecID

AdoQry.Parameters.Append AdoParameter1
AdoQry.Parameters.Append AdoParameter2


On Error GoTo modErr1
If Len(Dir$(sFileName)) > 0 Then Kill sFileName

DataFile = FreeFile
Open sFileName For Binary Access Write As DataFile

'With rdoQy
'    Set .ActiveConnection = RdoCon
'    .sql = "{? = CALL Qry_ImageData(?) }"
'End With

'rdoQy(0).Direction = rdParamReturnValue
'rdoQy(1).Direction = rdParamInput
'rdoQy(1).Value = iRecID
'
bSqlRows = clsADOCon.GetQuerySet(ADOPic, AdoQry, ES_KEYSET)
'Set RdoPic = rdoQy.OpenResultset(rdOpenKeyset, rdConcurRowVer)

'If (RdoPic.EOF = True) And (RdoPic.BOF = True) Then
If Not bSqlRows Then
    MsgBox "There Is No Logo in the database.", _
            vbInformation, "No Logo"
    ' Close the query
    Close DataFile
    'recordset
    'RdoPic.Close
    Set ADOPic = Nothing
    ReadImageFromDB = False
    Exit Function
End If

'GetChunk IMAGE Column

lngImgLen = ADOPic.Fields(1).ActualSize
intBlocks = lngImgLen \ lngBlockSize
intBlocksLo = lngImgLen Mod lngBlockSize

ReDim bytData(intBlocksLo)
bytData() = ADOPic.Fields(1).GetChunk(intBlocksLo)
Put DataFile, , bytData()

For intCnt = 1 To intBlocks
    bytData() = ADOPic.Fields(1).GetChunk(lngBlockSize)
    Put DataFile, , bytData()
Next intCnt

Close DataFile

' Close the recordset
'RdoPic.Close
' If every thing is good
ReadImageFromDB = True

 Exit Function
modErr1:
   sProcName = "ReadImageFromDB"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function


Public Function CreateDirectory(ByVal sFilePath As String) As Long
    Dim lRet As Long
    Dim sTemplateDir As String
    sTemplateDir = "C:\Windows"
    lRet = CreateDirectoryEx(sTemplateDir, sFilePath, ByVal 0&)
    CreateDirectory = lRet
End Function

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

Public Function String2Currency(str As String)
   If IsNumeric(str) Then
      String2Currency = CCur(str)
   Else
      String2Currency = CCur("0" & str)
   End If
End Function

Public Sub ClickedOnLogo()
    OpenWebPage "http://www.fusionerp.net/"
End Sub

Public Function GetBookPrice(strPart As String, strBook As String, ByRef strPrice As String)
   Dim RdoBok As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
          & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND " _
          & "(PBIREF='" & strBook & "') AND (PARTREF='" & Compress(strPart) & "')"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
   If bSqlRows Then
      With RdoBok
         strPrice = Format(!PBIPRICE, ES_SellingPriceFormat)
         ClearResultSet RdoBok
      End With
   Else
      strPrice = Format(0, ES_SellingPriceFormat)
   End If
   
   Set RdoBok = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function


Public Sub FillPartCombo(cmbPrt As ComboBox, Optional SelectFirst As Boolean = True)
   'Dim rdoPart As rdoResultset
   Dim rdoPart As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "Qry_FillParts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPart, ES_FORWARD)
   
   If bSqlRows Then
      With rdoPart
         While Not .EOF
            AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoPart = Nothing
   If SelectFirst Then
      cmbPrt.ListIndex = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillPartCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Sub

'Public Function RegisterSqlDsn(sDataSource As String) As String
'   Dim sAttribs As String
'   If sDataSource = "" Then sDataSource = "ESI2000"
'   sAttribs = "Description=" _
'              & "ES/2000ERP SQL Server Data " _
'              & vbCr & "OemToAnsi=No" _
'              & vbCr & "SERVER=" & sserver _
'              & vbCr & "Database=" & sDataBase
'   'Create new DSN or revise registered DSN.
'   rdoEngine.rdoRegisterDataSource sDataSource, _
'      "SQL Server", True, sAttribs
'   RegisterSqlDsn = sDataSource
'   Exit Function
'
'modErr1:
'   On Error GoTo 0
'   RegisterSqlDsn = sDataSource
'
'End Function

Public Function AddLog(nFileNum As Integer, strMsg As String)
   On Error GoTo DiaErr1
   
   If EOF(nFileNum) Then
      Print #nFileNum, strMsg
   End If

   Exit Function
DiaErr1:
   sProcName = "AddLog"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
End Function

Public Function GetPartSearchOption() As Boolean

   Dim rdo As ADODB.Recordset
   sSql = "SELECT ISNULL(PARTSEARCHOP,0) from ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   
   If bSqlRows Then
      If rdo.Fields(0) = 1 Then
         GetPartSearchOption = True
      Else
         GetPartSearchOption = False
      End If
   End If
   Set rdo = Nothing

End Function

Public Function SetupQtyEnabled() As Boolean
   Dim rdoSetup As ADODB.Recordset
   Dim iSet As Integer
   
   SetupQtyEnabled = False
   sSql = "SELECT ISNULL(COBOMSETQTY, 0) COBOMSETQTY FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSetup, ES_FORWARD)
   If bSqlRows Then
        iSet = 0 & rdoSetup!COBOMSETQTY
       If iSet = 1 Then SetupQtyEnabled = True
   End If
   Set rdoSetup = Nothing
End Function
