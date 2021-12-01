Attribute VB_Name = "ESIPOM"
Option Explicit

'from ESIPROJ in other modules
Public sFilePath As String


' Const string data (Poor Man's Resource File)
' Mostly msgbox type content.
Public Const SYSMSG0 = "Press The PIN# Button And Enter Your Employee Number Followed By The ENTER Button"
'Public Const SYSMSG1 = "Invalid Employee PIN" & vbCrLf & "Please Reenter."
Public Const SYSMSG1 = "Please reenter employee PIN"
Public Const SYSMSG2 = "Nothing To Do"
Public Const SYSMSG3 = "The Following Work Centers Are Available In Shop "
Public Const SYSMSG4 = "Select A Shop"
Public Const SYSMSG5 = "The Following Operations Are Available In Work Center "
Public Const SYSMSG6 = "Select A Job To Log Off."
Public Const SYSMSG7 = "System Error:  Cannot Log Off Job."
Public Const SYSMSG8 = "GetServerDateTime Error Has Occured."
Public Const SYSMSG9 = "No Operations Found In Workcenter."
Public Const SYSMSG10 = "Pick Items For Operation"
Public Const SYSMSG11 = "No Pick Items Found."
Public Const SYSMSG12 = "Available Lots For "
Public Const SYSMSG13 = "No Open Time Charges Journal For Current Date."
Public Const SYSMSG14 = "Inactive PIN. Please Reenter Active employee PIN"

Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Const CB_ADDSTRING = &H143

'************************************************************************************

Public gblnSqlRows As Boolean

Public Const MAX_CONCURRENT_LOGINS = 300
Public sInitials As String       ' Esi2000 login
Public bInvEmpFlag  As Boolean
Public ES_SYSDATE As Variant

' ES/2000 Company Info Datatype
Type CompanyInfo
   Name As String
   Addr(5) As String
   Phone As String
   Fax As String
   GlVerify As Byte
End Type

'************************************************************************************

' Menu Constants
Public Const MF_BYPOSITION = &H400&
Public Const MF_GRAYED = &H1&
Public Const SC_CLOSE = &HF060
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const SC_MOVE = &HF010
Public Const SC_RESTORE = &HF120

' Form Constants
Public Const ES_RESIZE = 0
Public Const ES_DONTRESIZE = -1
Public Const ES_LIST = 0
Public Const ES_DONTLIST = -1
Public Const ES_IGNOREDASHES = 1 'Compress routine

' MsgBox
Public Const ES_NOQUESTION = &H124 'Question and return (Default NO)
Public Const ES_YESQUESTION = &H24 'Question and return (Default YES)

' StrCase funtion contstants
Public Const ES_FIRSTWORD As Byte = 1

' Cursor types
Public Const ES_FORWARD = 0 'Default
Public Const ES_KEYSET = 1
Public Const ES_DYNAMIC = 2
Public Const ES_STATIC = 3

' SetWindowPos Support
Public Const Swp_NOMOVE = 2
Public Const Swp_NOSIZE = 1
Public Const Flags = Swp_NOMOVE Or Swp_NOSIZE
Public Const hWnd_TopMost = -1
Public Const Hwnd_NOTOPMOST = -2

' Colors
Global Const YELLOW = &HFFFF&

' Grid
Global Const ROWSPERPAGE = 8

'******************************** Global Varibles ************************************

' RDO
'Public RdoCon As rdoConnection
'Public RdoEnv As rdoEnvironment
'Public RdoErr As rdoError
'Public RdoRes As ADODB.Recordset

' ADO
Public clsADOCon As ClassFusionADO

' Global Project Varibles
Global glblActive As Label
Global gstrFilePath As String
Global gstrUser As String
Global gstrPassword As String
'Global gstrDatabase     As String
'Global gstrSaAdmin      As String
'Global gstrSaPassword   As String
'Global gstrServer       As String
Global sSql As String
Global gstrFacility As String
Global gstrJournal As String

Global gblnUserAction As Boolean
'Global gblnSqlRows      As Boolean

Global gbytScreen As Byte ' What screen we are on
Global Const LOGIN = &H1
Global Const SHOPS = &H2
Global Const WCS = &H3
Global Const jobs = &H4
Global Const complete = &H5
Global Const PKLIST = &H6
Global Const Lots = &H7


'Global Co               As CompanyInfo
'Global CurrError        As ModuleErrors

Global gbytMatrix() As Byte
Global gstrCaption As String
Global gstrCurRoutine As String ' Current Sub Or Function We Are In

Global gintResponse As Integer ' User response from frmAlert (vbyes or vbno)

Global gsngOverHead As Single ' Global Overhead


Global gblnTime As Boolean ' Enable display time charges
Public bSqlRows As Boolean

Public sProcName As String

'**********************************************************************************

' Windows API
Declare Function SetWindowPos Lib "user32" _
   (ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

'**********************************************************************************

Public Function SetTopMostWindow(plngHwnd As Long, _
                                 pblnTopmost As Boolean) As Long
   ' Make the window topmost
   If pblnTopmost Then
      SetTopMostWindow = SetWindowPos(plngHwnd, hWnd_TopMost, 0, 0, 0, 0, Flags)
   Else
      SetTopMostWindow = SetWindowPos(plngHwnd, Hwnd_NOTOPMOST, 0, 0, 0, 0, Flags)
      SetTopMostWindow = False
   End If
End Function

Public Function GetOverheadRate(mintEmpNum As Long, mstrPart As String, mlngRun As Long, mintOp As Integer, msngEmpRate As Single) As Single
   Dim mrdoOverHead As ADODB.Recordset
   sSql = "SELECT RnopTable.OPREF, RnopTable.OPRUN, RnopTable.OPNO, RnopTable.OPSHOP, " _
          & " RnopTable.OPCENTER, WcntTable.WCNOHFIXED, WcntTable.WCNOHPCT , WcntTable.WCNSTDRATE " _
          & " FROM RnopTable INNER JOIN WcntTable ON RnopTable.OPSHOP = WcntTable.WCNSHOP AND " _
          & " RnopTable.OPCENTER = WcntTable.WCNREF " _
          & "WHERE (RnopTable.OPREF = '" & Compress(mstrPart) & "') AND (RnopTable.OPRUN = " & mlngRun & ") AND (RnopTable.OPNO = " & mintOp & ")"
   
   
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, mrdoOverHead)
   If gblnSqlRows Then
      With mrdoOverHead
         
         '//////////////// NOT THIS WAY ANYMORE /////////////////
         'If !WCNOHPCT <> 0 Then
         'GetOverheadRate = Format(((!WCNOHPCT / 100) * msngEmpRate), "0.00")
         'Else
         'GetOverheadRate = Format(!WCNOHFIXED, "0.00")
         'End If
         '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
         'GetOverheadRate = Format((!WCNOHPCT / 100), "0.00000")
         GetOverheadRate = !WCNOHPCT
      End With
   Else
      GetOverheadRate = 0
   End If
End Function


'Public Function GetOpenJournal(sJrDate As String) As String
'   Dim rdoJrn As ADODB.Recordset
'   sSql = "SELECT COREF,COGLVERIFY FROM ComnTable WHERE COREF=1"
'   gblnSqlRows = clsADOCon.GetDataSet(sSql,rdoJrn, ES_FORWARD)
'   If gblnSqlRows Then Co.GlVerify = rdoJrn!COGLVERIFY
'
'   ' I'm taking this out for now
'   ' Co.GlVerify means to verify accounts of sub journals
'   ' when rolling up into GL
'
'   'If Co.GlVerify = 1 Then
'   'See if there is one
'   sSql = "SELECT MJTYPE,MJSTART,MJEND,MJCLOSED,MJGLJRNL FROM JrhdTable " _
'          & "WHERE MJTYPE='TJ' AND ('" & sJrDate & "' " _
'          & "BETWEEN MJSTART AND MJEND) AND MJCLOSED IS NULL"
'   gblnSqlRows = clsADOCon.GetDataSet(sSql,rdoJrn, ES_FORWARD)
'   If gblnSqlRows Then
'      With rdoJrn
'         If Not IsNull(.Fields(4)) Then
'            GetOpenJournal = .Fields(4)
'         Else
'            GetOpenJournal = ""
'         End If
'         .Cancel
'      End With
'      '   End If
'      'Else
'      '    GetOpenJournal = "None Required"
'   End If
'
'
'End Function
'
'

Public Function GetEmployee(pintPIN As Long, ByRef pempGetEmployee As Employee) As Boolean
   
   Dim prdoEmpl As ADODB.Recordset, logins As String
   logins = ShowCurrentLogins(pintPIN)
   If logins <> "" Then
      logins = "Current login status:" & vbCrLf & logins & vbCrLf
   End If
   sSql = "SELECT PREMLSTNAME,PREMFSTNAME,PREMPAYRATE,PREMSTATUS," _
          & "PREMACCTS,PREMTERMDT,PREMREHIREDT FROM EmplTable " _
          & "Where PREMNUMBER = " & pintPIN
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoEmpl)
   If gblnSqlRows Then
      'pempGetEmployee.intNumber = pintPIN
      With prdoEmpl
         ' Exit if the employee is terminated
         If (Not IsNull(!PREMTERMDT) And IsNull(!PREMREHIREDT)) Then
            MsgBox "Not a Current Employee.", vbCritical
            Set prdoEmpl = Nothing
            GetEmployee = False
            Exit Function
         ElseIf (Not IsNull(!PREMTERMDT) And Not IsNull(!PREMREHIREDT)) Then
            If (CDate(!PREMTERMDT) > CDate(!PREMREHIREDT)) Then
               MsgBox "Not a Current Employee.", vbCritical
               Set prdoEmpl = Nothing
               GetEmployee = False
               Exit Function
            End If
         End If
                  
        If (!PREMSTATUS) = "I" Then
           bInvEmpFlag = True
           Set prdoEmpl = Nothing
           GetEmployee = False
           Exit Function
         End If

         Select Case MsgBox(logins & "Do you wish to log in as " _
            & Trim(!PREMFSTNAME) & " " & Trim(!PREMLSTNAME) & " (" & pintPIN & ")?", vbSystemModal + vbYesNo)
         Case vbYes
            If Trim(!PREMACCTS) = "" Then
               MsgBox "There is no account defined for you.  You cannot log in.", vbCritical
               Exit Function
            End If
            Debug.Print "Login confirmed."
         Case Else
            Exit Function
         End Select
         
         pempGetEmployee.intNumber = pintPIN
         pempGetEmployee.strFirstName = "" & Trim(!PREMFSTNAME)
         pempGetEmployee.strLastName = "" & Trim(!PREMLSTNAME)
         pempGetEmployee.sngRate = !PREMPAYRATE
         pempGetEmployee.strAccount = "" & Trim(!PREMACCTS)
      End With
      If gblnTime Then
         'pempGetEmployee.strTimeCard = GetCardNumber(pempGetEmployee, Now)
         Dim tc As New ClassTimeCharge
         pempGetEmployee.strTimeCard = tc.GetTimeCardID(pempGetEmployee.intNumber, Now)
      End If
      PunchIn pempGetEmployee
      
      GetEmployee = True
   End If
End Function

Public Sub SystemAlert( _
                       pstrMsg As String, _
                       Optional pstrCaption As String, _
                       Optional pbln_Just_A_Message As Boolean, _
                       Optional pblnYesNo As Boolean)
   
   Dim plngTopWindow As Long
   
'   Unload frmAlert
'   frmAlert.Show vbModal
   With frmAlert
      
      .lblMsg = pstrMsg
      
      If pstrCaption <> "" Then
         .Caption = pstrCaption
      Else
         .Caption = frmMain.Caption
      End If
      
      plngTopWindow = SetTopMostWindow(.hwnd, True)
      
      If pbln_Just_A_Message Then
         .cmdOK.Visible = False
         .imgHalt.Visible = False
         .imgQuestion.Visible = False
         .imgInfo.Visible = True
      Else
         .imgInfo.Visible = False
         If pblnYesNo Then
            .imgQuestion.Visible = True
            .cmdOK.Visible = False
            .imgHalt.Visible = False
         Else
            .imgQuestion.Visible = False
            .cmdOK.Visible = True
            .imgHalt.Visible = True
         End If
      End If
      
      .cmdYes.Visible = pblnYesNo
      .cmdNo.Visible = pblnYesNo
      
      .Show vbModal, frmMain
      
   End With
   plngTopWindow = SetTopMostWindow(frm12Key.hwnd, True)
End Sub

Public Sub Main()
    Dim intResponse As Integer
    Dim sMsg As String
    Dim bRegOK As Boolean
    
   frmMain.Show
   
   '*
   '* This will need to change (nth)
   '* gstrUser and gstrPassword
   '*
   
   gstrUser = "ShopUser"
   gstrPassword = "ShopUser"
   
   'get Esi2000 user who last logged in on this machine
   sInitials = Trim(GetSetting("Esi2000", "System", "UserInitials", sInitials))
   
   '    gstrServer = UCase$( _
   '        GetSetting("Esi2000", _
   '        "System", "ServerId", _
   '        gstrServer))
   
   gstrFilePath = _
                  GetSetting("Esi2000", _
                  "System", "FilePath", _
                  gstrFilePath)
   
   RegisterUser
   OpenDBServer
   ' MM RDO comment
   'OpenSqlServer
   'UpdateTables 'the old scheme - remove it eventually  8/11/08
   ' MM RDO comment
   UpdateDatabase 'the new scheme
   

   'MM We should not be using CR 8.5
   'GetCrystalDSN
   
   
   ' get settings / just time for now
   Dim rdoTim As ADODB.Recordset
   sSql = "SELECT COPOMTIME From ComnTable"
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdoTim)
   If gblnSqlRows Then
      With rdoTim
         gblnTime = CBool("0" & .Fields(0))
         .Cancel
      End With
   End If
   Set rdoTim = Nothing
End Sub

Public Function CheckPath(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        CheckPath = True
    Else
        CheckPath = False
    End If
End Function

'Public Function GetSysLogon(pblnSaPassword As Boolean) As String
'    Dim a           As Integer
'    Dim i           As Integer
'    Dim sTest       As String
'    Dim sNewString  As String
'    Dim sPassword   As String
'
'    If pblnSaPassword Then
'        GetSysLogon = GetSetting("UserObjects", "System", "NoReg", GetSysLogon)
'        If Trim(GetSysLogon) = "" Then GetSysLogon = "sa"
'    Else
'        sPassword = GetSetting("SysCan", "System", "RegOne", sPassword)
'        sPassword = Trim(sPassword)
'        If sPassword <> "" Then
'            i = Len(sPassword)
'            If i > 5 Then
'                sPassword = Mid(sPassword, 4, i - 5)
'            End If
'        End If
'        GetSysLogon = sPassword
'    End If
'End Function
'
'Standard procedure for receiving resultsets
'bSqlRows = GetDataSet (RdoRes, ES_FORWARD)
'See also GetQuerySet for VB built queries

Sub MouseCursor(MCursor As Integer)
   'Allows consistant MousePointer Updates
   Screen.MousePointer = MCursor
   gblnUserAction = True
End Sub

Sub GetCompany(Optional bWantAddress As Byte)
   Dim ActRs As ADODB.Recordset
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
   gblnSqlRows = clsADOCon.GetDataSet(sSql, ActRs, ES_STATIC)
   If gblnSqlRows Then
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
   gstrFacility = Co.Name
   Set ActRs = Nothing
   Exit Sub
   
modErr1:
   Resume modErr2
modErr2:
   On Error GoTo 0
End Sub

Public Function GetSystemMessage() As String
   Static pstrOldMessage As String
   Dim pRdoMsg As ADODB.Recordset
   
   On Error GoTo modErr1
   
   'sSql = "Qry_GetSysMessage"
   sSql = "select ALERTMSG from Alerts where ALERTREF = 1"
   Set pRdoMsg = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
   
   If Not pRdoMsg.BOF And Not pRdoMsg.EOF Then
      With pRdoMsg
'         If pstrOldMessage <> "" & Trim(!ALERTMSG) Then
'            pstrOldMessage = "" & Trim(!ALERTMSG)
'            GetSystemMessage = pstrOldMessage
'         Else
'            GetSystemMessage = ""
'         End If
         GetSystemMessage = "" & Trim(!ALERTMSG)
      End With
   End If
   
   Set pRdoMsg = Nothing
   
   Exit Function
   
modErr1:
   On Error GoTo 0
End Function

Public Sub RegisterUser()
   Dim pintFileNumber As Integer
   
   'pintFileNumber = FreeFile
   'Open gstrFilePath & "\ES_USERS.LOG" For Append As #pintFileNumber
   '    Print #pintFileNumber, "Worktstation       User            Time"
   'Close #pintFileNumber
End Sub

' Old ES/200X source compress routine

Public Function Compress( _
                         TestNo As Variant, _
                         Optional iLength As Integer, _
                         Optional bIgnoreDashes As Byte) As String
   
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
   Resume modErr2
modErr2:
   On Error Resume Next
   Compress = TestNo
End Function

Public Sub CenterPictureBox( _
                            ppicToCenter As PictureBox, _
                            pfrmToCenterIn As Form)
   
   ppicToCenter.Left = (pfrmToCenterIn.Width - ppicToCenter.Width) / 2
   
   ' Note: 200 is to offset the menu bar
   ppicToCenter.Top = ((pfrmToCenterIn.Height - ppicToCenter.Height) / 2) - 500
End Sub

' Return Morning, Afternoon, or Evening

Public Function Time_Of_Day() As String
   If time < TimeValue("12:00") And time >= TimeValue("00:00") Then
      Time_Of_Day = "Morning"
   ElseIf time > TimeValue("11:59") And time < TimeValue("18:00") Then
      Time_Of_Day = "Afternoon"
   ElseIf time >= TimeValue("18:00") And time <= TimeValue("23:59") Then
      Time_Of_Day = "Evening"
   End If
End Function


'''Public Function GetCardNumber(pempEmployee As Employee, LoginTime As Date) As String
'''   'Provides for a unique time card number
'''   'Stored as an (11) char string
'''   'The First (5) is the card date
'''   'the last (6) pickup the time to the part of a
'''   'second.
'''
'''   Dim s As Single
'''   Dim l As Long
'''   Dim m As Long
'''   Dim T As String
'''   Dim prdoRes As ADODB.Recordset
'''   Dim pstrTmCard As String
'''   Dim loginDate As String
'''   loginDate = Format(LoginTime, "mm/dd/yyyy")
'''
'''   ' Check for existing time card for the day, if found use it for all
'''   ' other time charges to jobs for the day.
''''   sSql = "SELECT TMCARD FROM TchdTable WHERE " _
''''          & "TMEMP=" & pempEmployee.intNumber _
''''          & " AND TMDAY='" & Format(GetServerDateTime, "mm/dd/yyyy") & "'"
'''   sSql = "SELECT TMCARD FROM TchdTable WHERE " _
'''          & "TMEMP=" & pempEmployee.intNumber _
'''          & " AND TMDAY='" & loginDate & "'"
'''   gblnSqlRows = clsADOCon.GetDataSet(sSql,prdoRes)
'''
'''   If gblnSqlRows Then
'''      With prdoRes
'''         pstrTmCard = Trim(.Fields(0))
'''      End With
'''   Else
'''      m = DateValue(Format(GetServerDateTime, "yyyy,mm,dd"))
'''      s = TimeValue(Format(GetServerDateTime, "hh:mm:ss"))
'''      l = s * 1000000
'''      pstrTmCard = Format(m, "00000") & Format(l, "000000")
'''      sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDAY) VALUES ('" _
'''             & pstrTmCard & "'," & pempEmployee.intNumber & ",'" _
'''             & Format(GetServerDateTime, "mm/dd/yyyy") & "')"
'''      clsADOCon.ExecuteSQL sSql ' rdExecDirect
'''   End If
'''   Set prdoRes = Nothing
'''   GetCardNumber = pstrTmCard
'''End Function
'''
' Return jobs employee is logged into...

Public Sub GetCurrentJobs(pintEmpl As Long, ByRef pjobMyJobs() As Job)
   Dim prdoJobs As ADODB.Recordset
   Dim jobs As Integer
   
   ReDim pjobMyJobs(0)
   
   sSql = "SELECT PARTNUM, ISRUN, ISOP FROM IstcTable INNER JOIN " _
          & "PartTable ON IstcTable.ISMO = PartTable.PARTREF " _
          & "Where ISEMPLOYEE = " & pintEmpl & " and ReadyToDelete = 0 and ISINDIRECT = 0"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoJobs)
   If gblnSqlRows Then
      With prdoJobs
         While Not .EOF
            ReDim Preserve pjobMyJobs(jobs)
            pjobMyJobs(jobs).strPart = Trim(.Fields(0))
            pjobMyJobs(jobs).lngRun = .Fields(1)
            pjobMyJobs(jobs).intOp = .Fields(2)
            'pjobMyJobs(jobs).strPart = .Fields(0)
            .MoveNext
            jobs = jobs + 1
         Wend
      End With
   End If
   Set prdoJobs = Nothing
   'MsgBox ShowCurrentLogins(pintEmpl)
End Sub

Public Function ShowCurrentLogins(empno As Long) As String
   Dim rdo As ADODB.Recordset
   Dim msg
  
   sSql = "SELECT PARTNUM, ISRUN, ISOP, ISINDIRECT, ISMOSTART, ISOPLOGOFF FROM IstcTable" & vbCrLf _
      & "LEFT JOIN PartTable ON IstcTable.ISMO = PartTable.PARTREF " & vbCrLf _
      & "Where ISEMPLOYEE = " & empno & " and ReadyToDelete = 0"
   
   If clsADOCon.GetDataSet(sSql, rdo) Then
      With rdo
         While Not .EOF
            If !ISINDIRECT = 0 Then
               If !ISOPLOGOFF = 0 Then
                  'msg = msg & "LOGGED IN:   "
               Else
                  msg = msg & "PUNCHOUT: "
               End If
               msg = msg & Trim(!PartNum) & " run " & !ISRUN & " op " & !ISOP
               If !ISOPLOGOFF = 0 Then
                  msg = msg & " " & Format(!ISMOSTART, "MM/DD hh:mm AM/PM")
               End If
            Else
               msg = msg & "INDIRECT: " _
                  & Format(!ISMOSTART, "MM/DD hh:mm AM/PM")
            End If
            msg = msg & vbCrLf
            
            'indicate if charge will be truncated as 16 hours
            If !ISOPLOGOFF = 0 Then
               If DateDiff("h", !ISMOSTART, Now) >= 16 Then
                  msg = msg & "**** WILL AUTO LOGOFF AT 16 HOURS" & vbCrLf
               End If
            End If
            
            .MoveNext
         Wend
      End With
   End If
   ShowCurrentLogins = msg

End Function


Public Sub OpenIndirectTC(pempEmployee As Employee)
   Dim prdoExists As ADODB.Recordset
   
   If Not gblnTime Then Exit Sub
   
   'don't create an indirect charge if any current charge, direct or indirect, exists for employee
   sSql = "SELECT ISMOSTART FROM IstcTable" & vbCrLf _
      & "WHERE ISEMPLOYEE = " & pempEmployee.intNumber
   If clsADOCon.GetDataSet(sSql, prdoExists) = 0 Then
      sSql = "INSERT INTO IstcTable (ISEMPLOYEE, ISMO, ISRUN,ISOP, ISMOSTART, ISINDIRECT)" & vbCrLf _
         & "VALUES(" & pempEmployee.intNumber & ",'', 0, 0, " & vbCrLf _
         & "'" & GetServerDateTimeToMinute & "', 1)"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   End If
End Sub


Public Sub CloseIndirectTC(pempEmployee As Employee)
   ' Close an indirect time charge (if any) usually occurs when first job is logged into.
   Dim pdteStart As Date
   Dim pdteStop As Date
   Dim psngHrs As Single
   Dim pstrEnd As String
   Dim prdoStart As ADODB.Recordset
   Dim pdteTCTIME As Date
   Dim rdo As ADODB.Recordset
   
   If Not gblnTime Then Exit Sub
   
   On Error GoTo whoops
   
   'find a regular timecode
   Dim regularTimeCode As String
   sSql = "select TYPECODE from TmcdTable where TYPETYPE = 'R' ORDER BY TYPESEQ"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      regularTimeCode = rdo.Fields(0)
   Else
      regularTimeCode = "RT"
   End If
   Set rdo = Nothing
   
   Err.Clear
   
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   'get the starting time for the indirect time charge
   Dim startTime As Date
   sSql = "SELECT ISMOSTART FROM IstcTable" & vbCrLf _
          & "WHERE ISEMPLOYEE = " & pempEmployee.intNumber & vbCrLf _
          & "AND ISINDIRECT = 1"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      startTime = rdo.Fields(0)
   Else
      clsADOCon.RollbackTrans
      clsADOCon.ADOErrNum = 0
      Set rdo = Nothing
      Exit Sub
   End If
   
   Set rdo = Nothing
   
   'determine ending time - limit to 16 hours
   Dim endTime As Date
   endTime = GetServerDateTimeToMinute 'Now
   If DateDiff("n", startTime, endTime) > 16 * 60 Then
      endTime = DateAdd("n", 16 * 60, startTime)
   End If
   
   'get the timecard
   Dim tc As New ClassTimeCharge
   Dim tcid As String
   Dim strSftBegTime As String
   
   strSftBegTime = tc.GetShiftStartDate(pempEmployee.intNumber, startTime)
   
   tcid = tc.GetTimeCardID(pempEmployee.intNumber, startTime, strSftBegTime)

   'create the time charge
   tc.CreateTimeCharge tcid, pempEmployee.intNumber, startTime, endTime, _
      regularTimeCode, "I", pempEmployee.strAccount, "", 0, 0, 0, "POM", 0, 0, 0, ""
      
   If (strSftBegTime <> "") Then
      tc.UpdateTimeCardTotals pempEmployee.intNumber, CDate(strSftBegTime), True
   Else
      tc.UpdateTimeCardTotals pempEmployee.intNumber, startTime, True
   End If
   
   'tc.ComputeOverlappingCharges pempEmployee.intNumber, startTime   'not for indirect
   
   'tc.UpdateTimeCardTotals pempEmployee.intNumber, startTime, True  'seems redundant and may cause error
   
   'delete temp indirect charge
   sSql = "delete from IstcTable where ISEMPLOYEE = " & pempEmployee.intNumber & " and ISINDIRECT=1"
   clsADOCon.ExecuteSql sSql
   clsADOCon.CommitTrans
   
'   If Err Then
'      clsADOCon.RollbackTrans
'      MsgBox "Update failed: " & Err.Description, vbInformation, "CloseIndirectTC"
'   Else
'      clsADOCon.CommitTrans
'      'Now update the
'      'tc.UpdateTimeCardTotals pempEmployee.intNumber, startTime, True
'
'   End If
   
   Exit Sub
   
whoops:
   sProcName = "CloseIndirectTC"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Sub


Public Function PunchOut(pempEmployee As Employee) As Boolean
   ' return = True if punchout confirmed
   '        = False if user canceled
   'confirm
   Select Case MsgBox("Do you wish to punch out " _
      & Trim(pempEmployee.strFirstName) & " " & Trim(pempEmployee.strLastName) _
      & " (" & pempEmployee.intNumber & ")?", vbSystemModal + vbYesNo)
   Case vbYes
      Debug.Print "Login confirmed."
   Case Else
      Exit Function
   End Select

   PunchOut = True
   
   Dim jobs As Integer
   CloseIndirectTC pempEmployee
   If pempEmployee.jobCurMO(0).lngRun <> 0 Then
      For jobs = 0 To UBound(pempEmployee.jobCurMO)
         LogOffJob _
            pempEmployee, _
            pempEmployee.jobCurMO(jobs), 0, 0, 0, False, False
      Next
   End If
End Function

Public Sub PunchIn(pempEmployee As Employee)
   ' Check for any jobs that are punched out but not complete.  If found log back in.
   
   Dim lRecAffected As Long
   On Error GoTo whoops
   ' MM clsADOCon.BeginTrans
   clsADOCon.BeginNestedTrans
   
   'if there are dlirect charge logins, there should be no indirect charges
   Dim rdoPunchIn As ADODB.Recordset
   sSql = "select * from IstcTable where ISINDIRECT = 0 and ISEMPLOYEE = " & pempEmployee.intNumber
'   If clsADOCon.GetDataSet(sSql, rdoPunchIn) Then
'      sSql = "delete from IstcTable where ISEMPLOYEE = " & pempEmployee.intNumber & vbCrLf _
'         & "AND ISINDIRECT = 1"
'      clsADOCon.ExecuteSql sSql, lRecAffected
'      If lRecAffected > 0 Then
'         MsgBox "Deleting overlapping indirect charge started " & Format(rdoPunchIn!ISMOSTART, "mm/dd/yy hh:mm AM/PM")
'      End If

   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPunchIn)
   If bSqlRows Then
      Dim dtString As String
      dtString = Format(rdoPunchIn!ISMOSTART, "mm/dd/yy hh:mm AM/PM")
      rdoPunchIn.Close
      
      sSql = "delete from IstcTable where ISEMPLOYEE = " & pempEmployee.intNumber & vbCrLf _
         & "AND ISINDIRECT = 1"
      clsADOCon.ExecuteSql sSql, lRecAffected
      If lRecAffected > 0 Then
         MsgBox "Deleting overlapping indirect charge started " & dtString
      End If
   
   'if there are no direct charges, there should be an indirect charge
   'if one already exists, none will be created
   Else
     rdoPunchIn.Close
     OpenIndirectTC pempEmployee
   End If
   
   Set rdoPunchIn = Nothing
   
   'create time charges for any indirect logins > 16 hours, truncating at 16 hours
   sSql = "select * from IstcTable where datediff(hh, ISMOSTART, getdate()) > 16" & vbCrLf _
      & "and ISEMPLOYEE = " & pempEmployee.intNumber & " and ISINDIRECT = 1"
   If clsADOCon.GetDataSet(sSql, rdoPunchIn) Then
      CloseIndirectTC pempEmployee
   End If
   
   Set rdoPunchIn = Nothing
   'create time charges for any direct logins > 16 hours, truncating at 16 hours
   sSql = "select * from IstcTable where datediff(hh, ISMOSTART, getdate()) > 16" & vbCrLf _
      & "and ISEMPLOYEE = " & pempEmployee.intNumber & " and ISINDIRECT = 0" & vbCrLf _
      & "and ISOPLOGOFF = 0"
   If clsADOCon.GetDataSet(sSql, rdoPunchIn, ES_STATIC) Then
      While Not rdoPunchIn.EOF
         Dim j As Job
         
         j.strPart = Trim(rdoPunchIn!ISMO)
         j.lngRun = rdoPunchIn!ISRUN
         j.intOp = rdoPunchIn!ISOP
         j.sngQty = 0
         LogOffJob pempEmployee, j, 0, 0, 0, False, False
         rdoPunchIn.MoveNext
      Wend
   End If
   
   Set rdoPunchIn = Nothing
   'log back to jobs punched out of
   sSql = "UPDATE IstcTable SET ISMOSTART = '" _
      & GetServerDateTimeToMinute & "',ISOPLOGOFF = 0 WHERE ISEMPLOYEE = " _
      & pempEmployee.intNumber & " AND ISOPLOGOFF = 1"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   'if there are no remaining direct charges, add an indirect charges
   sSql = "select * from IstcTable where ISINDIRECT = 0 and ISEMPLOYEE = " & pempEmployee.intNumber
   If clsADOCon.GetDataSet(sSql, rdoPunchIn) = 0 Then
      OpenIndirectTC pempEmployee
   End If
   Set rdoPunchIn = Nothing
   
   clsADOCon.CommitNestedTrans
   Exit Sub
   
whoops:
   clsADOCon.RollbackNestedTrans
   sProcName = "PunchIn"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   
End Sub

Private Function GetEmployeeAccount(pempEmployee As Employee, pjobJob As Job) As String
   Dim prdoEmpAcct As ADODB.Recordset
   sSql = "SELECT WCNACCT, SHPACCT " _
          & " FROM RnopTable INNER JOIN " _
          & " WcntTable ON RnopTable.OPCENTER = WcntTable.WCNREF INNER JOIN " _
          & " ShopTable ON RnopTable.OPSHOP = ShopTable.SHPREF " _
          & " WHERE (RnopTable.OPREF = '" & pjobJob.strPart & "') AND (RnopTable.OPRUN = " & pjobJob.lngRun _
          & " ) AND (RnopTable.OPNO = " & pjobJob.intOp & ")"
   
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoEmpAcct)
   If gblnSqlRows Then
      With prdoEmpAcct
         If "" & Trim(!WCNACCT) <> "" Then
            GetEmployeeAccount = Trim(!WCNACCT)
         Else
            If "" & Trim(!SHPACCT) <> "" Then
               GetEmployeeAccount = "" & Trim(!SHPACCT)
            Else
               GetEmployeeAccount = Trim(pempEmployee.strAccount)
            End If
         End If
      End With
   Else
      GetEmployeeAccount = Trim(pempEmployee.strAccount)
   End If
End Function



Public Sub LogOffJob( _
                     pempEmployee As Employee, _
                     pjobComplete As Job, _
                     psngAccept As Single, _
                     psngReject As Single, _
                     psngScrap As Single, _
                     IsOperationComplete As Boolean, _
                     IsUserLoggingOff As Boolean, _
                     Optional Comments As String, _
                     Optional bMOComt As Boolean = False)
   
   'log off of job or punch out
   'IsUserLoggingOff   = True if user is logging off of job
   '                   = False if user is punching out
   '                     (jobs are retained for next login)
   
   gstrCurRoutine = "LogOffJob"
   
   Dim TimeCardID As String
   
   Dim prdoStart As ADODB.Recordset
   Dim prdoJobs As ADODB.Recordset
   
   Dim LoginTime As Date ' Starting time
   Dim LogoutTime As Date ' Stopping time
   Dim psngHrs As Single ' hours allocated to job / may be prorated
   
   Dim pstrShop As String
   Dim pstrWcnt As String
   Dim pstrEmpAccnt As String ' For getemployeeaccount function
   Dim pstrSetupTime As String
   Dim bForceLogOut As Boolean
   
   On Error GoTo whoops
   
   If IsUserLoggingOff Then
      SystemAlert "Logging Off... " & vbCrLf & vbCrLf _
         & Trim(pjobComplete.strPart) _
         & "  Run:" & pjobComplete.lngRun _
         & "  Op:" & pjobComplete.intOp, , True
   End If
   
   ' Get the starting time for the job.
   ' Note all job information comes from modular
   ' user defined varible mjobCurrent.
   
   ' Time as runtime
   pstrSetupTime = "R"
   sSql = "SELECT ISMOSTART, ISSHOP, ISWCNT,IsNull(ISSURUN,'R')  FROM IstcTable " & vbCrLf _
          & "WHERE ISEMPLOYEE = " & pempEmployee.intNumber & " AND " & vbCrLf _
          & "ISMO = '" & Compress(pjobComplete.strPart) & "' AND " & vbCrLf _
          & "ISRUN = " & pjobComplete.lngRun & " AND " & vbCrLf _
          & "ISOP = " & pjobComplete.intOp
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoStart)
   If gblnSqlRows Then
      With prdoStart
         LoginTime = .Fields(0)
         pstrShop = "" & Trim(.Fields(1))
         pstrWcnt = "" & Trim(.Fields(2))
         
         pstrSetupTime = "" & Trim(.Fields(3))
         If (pstrSetupTime = "") Then pstrSetupTime = "R"
         
      End With
   End If
   Set prdoStart = Nothing
   DoEvents
   
   'calculate logout time
   'don't allow total flow time > 16 hours
   LogoutTime = GetServerDateTimeToMinute
   If LogoutTime <= LoginTime Then LogoutTime = (LoginTime + "0:01:00")
   If DateDiff("n", LoginTime, LogoutTime) > 16 * 60 Then
      LogoutTime = DateAdd("h", 16, LoginTime)
   End If
   
   ' only process query if time keeping is turned on
   'If gblnTime And ChargedHours > 0 Then
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   If gblnTime And LoginTime <> LogoutTime Then
      Dim tc As New ClassTimeCharge
      Dim strSftBegTime As String
      strSftBegTime = tc.GetShiftStartDate(pempEmployee.intNumber, LoginTime)
      
      TimeCardID = tc.GetTimeCardID(pempEmployee.intNumber, LoginTime, strSftBegTime)   'get timechard for login date
      
      'find a regular timecode
      Dim rdo As ADODB.Recordset
      Dim regularTimeCode As String
      sSql = "select TYPECODE from TmcdTable where TYPETYPE = 'R' ORDER BY TYPESEQ"
      If clsADOCon.GetDataSet(sSql, rdo) Then
         regularTimeCode = rdo.Fields(0)
      Else
         regularTimeCode = "RT"
      End If
      Set rdo = Nothing
      
      If tc.CreateTimeCharge(TimeCardID, pempEmployee.intNumber, LoginTime, LogoutTime, _
         regularTimeCode, pstrSetupTime, "", _
         pjobComplete.strPart, pjobComplete.lngRun, pjobComplete.intOp, 0, "POM", _
         psngAccept, psngReject, psngScrap, Comments, bMOComt) Then
         
            tc.ComputeOverlappingCharges CLng(pempEmployee.intNumber), LoginTime
            
            If (strSftBegTime <> "") Then
               tc.ComputeOverlappingCharges CLng(pempEmployee.intNumber), CDate(strSftBegTime)
               tc.UpdateTimeCardTotals pempEmployee.intNumber, CDate(strSftBegTime), True
            Else
               tc.ComputeOverlappingCharges CLng(pempEmployee.intNumber), LoginTime
               tc.UpdateTimeCardTotals pempEmployee.intNumber, LoginTime, True
            End If
            ' MM 9/8 clsADOCon.CommitTrans
      
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Unable to create time charge"
         Exit Sub
      End If
      
   End If
   Err.Clear
   
   ' if the force logoff is set in the company setting.
   bForceLogOut = GetAutoLogputOnPunchout
   
   If (IsUserLoggingOff = True) Or (bForceLogOut = True) Then
      'sSql = "UPDATE IstcTable set ISMOEND='" & LogoutTime & "', ReadyToDelete=1" & vbCrLf
      sSql = "delete from IstcTable" & vbCrLf
   Else
      ' Just mark as punched out not completed
      sSql = "UPDATE IstcTable SET ISOPLOGOFF = 1 "
   End If
   
   sSql = sSql _
          & "WHERE ISEMPLOYEE = " & pempEmployee.intNumber & vbCrLf _
          & "AND ISMO = '" & Compress(pjobComplete.strPart) & "' " & vbCrLf _
          & "AND ISRUN = " & pjobComplete.lngRun & vbCrLf _
          & "AND ISOP = " & pjobComplete.intOp & vbCrLf _
          & "AND ISMOEND IS NULL"
   clsADOCon.ExecuteSql sSql
   
'   'delete any temp charges for this employee, where no overlapping charges
'   'remain to be calculated
'   sSql = "delete from istctable" & vbCrLf _
'          & "where isemployee=" & pempEmployee.intNumber & vbCrLf _
'          & "and ReadyToDelete = 1" & vbCrLf _
'          & "and ISMOEND <= ISNULL((select min(ISMOSTART) from istctable" & vbCrLf _
'          & "where ISMOEND is null and ISOPLOGOFF = 0 and isemployee=" & pempEmployee.intNumber & "), " & vbCrLf _
'          & "getdate())"
'   clsADOCon.ExecuteSQL sSql
'   Err.Clear
   
   UpdateOpFromTimeCharges pjobComplete.strPart, pjobComplete.lngRun, _
         pjobComplete.intOp, IsOperationComplete
   clsADOCon.CommitTrans
   Exit Sub
   
whoops:
   clsADOCon.RollbackTrans
   sProcName = "LogOffJob"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
End Sub





Private Sub AssignToMatrix( _
                           pdteBeginning As Date, _
                           pintRow As Integer, _
                           pdteStart As Date, _
                           pdteStop As Date, _
                           pintMinutes As Integer)
   
   Dim jobs As Integer
   Dim pbytInsert As Byte
   
   Dim pdteCompareDate As Date
   
   For jobs = 0 To pintMinutes
      pdteCompareDate = DateAdd("n", jobs, pdteBeginning)
      If pdteCompareDate >= pdteStart And pdteCompareDate <= pdteStop Then
         pbytInsert = 1
      Else
         pbytInsert = 0
      End If
      gbytMatrix(pintRow, jobs) = pbytInsert 'subscript error here
   Next
End Sub

Public Sub DisplayError()
   Dim pstrMsg As String
   pstrMsg = "The Following Error Has Occured:" _
             & vbCrLf & vbCrLf & "Error # " & Err _
             & vbCrLf & "Description: " & Err.Description _
             & vbCrLf & "Routine: " & gstrCurRoutine _
             & vbCrLf & vbCrLf & _
             "Please Contact System Administrator Or ESI"
   MsgBox pstrMsg, vbExclamation, gstrCaption
   Err = 0
End Sub

Public Function IsPickOp(pjobCurrent As Job) As Boolean
   Dim prdoPck As ADODB.Recordset
   gstrCurRoutine = "IsPickOp"
   sSql = "SELECT OPPICKOP FROM RnopTable WHERE " _
          & "OPREF = '" & Compress(pjobCurrent.strPart) & "' AND " _
          & "OPRUN = " & pjobCurrent.lngRun & " AND " _
          & "OPNO = " & pjobCurrent.intOp
   gblnSqlRows = clsADOCon.GetDataSet(sSql, prdoPck)
   If gblnSqlRows Then
      With prdoPck
         If .Fields(0) = 1 Then
            IsPickOp = True
         End If
      End With
   End If
   Set prdoPck = Nothing
End Function



' PickItems
' Taken directly from ES/2000 Production (nth)
' Modifed to accept a grid and job and fill vItems internally
' before picking items.

Public Sub PickItems( _
                     pgrdItems As MSFlexGrid, _
                     pjobCurrent As Job)
   
   Dim RdoPck As ADODB.Recordset
   Dim i As Integer
   Dim a As Integer
   Dim bBadPick As Byte
   Dim bGoodPick As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sMoStatus As String
   Dim sMon As String
   Dim sMoPart As String * 31
   Dim sMoRun As String * 9
   Dim sNewDate As String
   Dim sNewPart As String
   Dim sNewRev As String
   Dim sComment As String
   Dim cQuantity As Currency
   Dim cCost As Currency
   
   Dim sDebitAcct As String
   Dim sCreditAcct As String
   
   Dim sSql As String
   Dim bSqlRows As Boolean
   
   ' New load vItems Array
   Dim vitems(300, 10) As Variant
   '0 = loc
   '1 = part
   '2 = compressed part
   '3 = stand cost
   '4 = desc
   '5 = rev
   '6 = planned
   '7 = actual
   '8 = wip
   '9 = complete?
   Dim iTotalItems As Integer
   Dim k As Integer
   
   
   On Error GoTo DiaErr1
   'On Error GoTo 0
   
   With pgrdItems
      For i = 0 To pgrdItems.Rows - 1
         .Row = i
         If .CellBackColor = YELLOW Then
            k = k + 1
            
            .Col = 0
            vitems(k, 1) = .Text
            vitems(k, 2) = Compress(.Text)
            
            .Col = 2
            vitems(k, 6) = .Text
            
            .Col = 4
            vitems(k, 7) = .Text
            
            .Col = 5
            If .Text = "No" Then
               vitems(k, 9) = 0
            Else
               vitems(k, 9) = 1
            End If
            
            .Col = 6
            vitems(k, 3) = CSng(.Text)
         End If
      Next
   End With
   iTotalItems = k
   
   
   ' Everything seems ok, so let's pick it
   
   MouseCursor ccHourglass
   sMon = pjobCurrent.strPart '(lblMon)
   sMoPart = pjobCurrent.strPart
   i = Len(Trim(str(pjobCurrent.lngRun)))
   i = 5 - i
   sMoRun = "RUN" & Space$(i) & pjobCurrent.lngRun
   
   For i = 1 To iTotalItems
      If Val(vitems(i, 7)) > 0 Then
         'Activity
         
         GetAccounts CStr(vitems(i, 2)), sDebitAcct, sCreditAcct
         
         cQuantity = Val(vitems(i, 7)) - (Val(vitems(i, 7)) * 2)
         cCost = cCost + Val(vitems(i, 3))
         
         sSql = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2,INPQTY,INAQTY,INAMT," _
                & "INCREDITACCT,INDEBITACCT,INMOPART,INMORUN)  " _
                & "VALUES(10,'" & vitems(i, 2) & "','PICK','" & sMoPart & sMoRun & "'," _
                & vitems(i, 7) & "," & str(cQuantity) & "," & vitems(i, 3) & ",'" _
                & sCreditAcct & "','" & sDebitAcct & "','" _
                & sMon & "'," & pjobCurrent.lngRun & ")"
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         sNewPart = vitems(i, 2)
         'AverageCost sNewPart
         
         'pick
         sSql = "UPDATE MopkTable SET PKAQTY=" & vitems(i, 7) & ",PKTYPE=10," _
                & "PKAMT=" & vitems(i, 3) & ", PKADATE='" & Format(GetServerDateTime, "mm/dd/yy") & "'," _
                & "PKWIP='" & vitems(i, 8) & "' WHERE PKPARTREF='" & vitems(i, 2) & "' " _
                & "AND PKMOPART='" & pjobCurrent.strPart & "' AND PKMORUN=" & pjobCurrent.lngRun & " " _
                & "AND PKRECORD=" & Trim(str(i)) & " "
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
         sSql = "UPDATE PartTable SET PAQOH=PAQOH-" & Trim(str(Abs(cQuantity))) _
                & " WHERE PARTREF='" & vitems(i, 2) & "' "
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
         'See if we need to add one
         If Val(vitems(i, 7)) < Val(vitems(i, 6)) Then
            If Val(vitems(i, 9)) = 0 Then
               sMoStatus = "PP"
               If Trim(vitems(i, 5)) = "" Then
                  sNewRev = "A"
               Else
                  a = Asc(Left(vitems(i, 5), 1)) + 1
                  sNewRev = Chr(a)
               End If
               
               sSql = "SELECT * FROM MopkTable WHERE PKRECORD=" & Trim(str(i)) & " "
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_KEYSET)
               
               If bSqlRows Then
                  cQuantity = RdoPck!PKBOMQTY
                  a = RdoPck!PKMORUNOP
                  sMsg = "" & RdoPck!PKREFERENCE
                  sComment = "" & RdoPck!PKCOMT
                  sNewDate = Format$(RdoPck!PKPDATE, "mm/dd/yy")
                  RdoPck.AddNew
                  RdoPck!PKPARTREF = vitems(i, 2)
                  RdoPck!PKMOPART = pjobCurrent.strPart
                  RdoPck!PKMORUN = pjobCurrent.lngRun
                  RdoPck!PKMORUNOP = a
                  RdoPck!PKTYPE = 9
                  RdoPck!PKREV = sNewRev
                  RdoPck!PKPDATE = sNewDate
                  RdoPck!PKPQTY = Val(vitems(i, 6)) - Val(vitems(i, 7))
                  RdoPck!PKORIGQTY = Val(vitems(i, 6))
                  RdoPck!PKBOMQTY = cQuantity
                  RdoPck!PKREFERENCE = sMsg
                  RdoPck!PKCOMT = sComment
                  RdoPck.Update
               End If
            End If
         End If
      End If
   Next
   sSql = "UPDATE RunsTable SET RUNSTATUS='" & sMoStatus & "'," _
          & "RUNCMATL=RUNCMATL+" & Val(vitems(i, 3)) & "," _
          & "RUNCOST=RUNCOST+" & cCost & " " _
          & "WHERE RUNREF='" & pjobCurrent.strPart & "' AND RUNNO=" & pjobCurrent.lngRun & " "
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   sSql = "UPDATE MopkTable SET PKRECORD=0 WHERE PKRECORD>0 " _
          & "AND PKPARTREF='" & pjobCurrent.strPart & "' AND PKMORUN=" _
          & pjobCurrent.lngRun & " "
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   DoEvents
   On Error Resume Next
   Set RdoPck = Nothing
   MouseCursor ccArrow
   'MsgBox "Pick Completed Successfully.", vbInformation, Caption
   'Unload Me
   Exit Sub
   
DiaErr1:
   gstrCurRoutine = "PickItems"
   DisplayError
End Sub

'7/8/99

Public Sub GetAccounts( _
                       sPart As String, _
                       sDebitAcct As String, _
                       sCreditAcct As String)
   
   Dim RdoAct As ADODB.Recordset
   Dim bType As Byte
   Dim sPcode As String
   Dim sSql As String
   Dim bSqlRows As Boolean
   
   
   On Error GoTo DiaErr1
   'Use current Part
   sSql = "Qry_GetExtPartAccounts '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
   If bSqlRows Then
      With RdoAct
         sPcode = "" & Trim(!PAPRODCODE)
         bType = Format(!PALEVEL, "0")
         If bType = 6 Or bType = 7 Then
            sDebitAcct = "" & Trim(!PACGSEXPACCT)
            sCreditAcct = "" & Trim(!PAINVEXPACCT)
         Else
            sDebitAcct = "" & Trim(!PACGSMATACCT)
            sCreditAcct = "" & Trim(!PAINVMATACCT)
         End If
         .Cancel
      End With
   Else
      sCreditAcct = ""
      sDebitAcct = ""
      Exit Sub
   End If
   If sDebitAcct = "" Or sCreditAcct = "" Then
      'None in one or both there, try Product code
      sSql = "Qry_GetPCodeAccounts '" & sPcode & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
      If bSqlRows Then
         With RdoAct
            If bType = 6 Or bType = 7 Then
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSEXPACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVEXPACCT)
            Else
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(!PCCGSMATACCT)
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(!PCINVMATACCT)
            End If
            .Cancel
         End With
      End If
      If sDebitAcct = "" Or sCreditAcct = "" Then
         'Still none, we'll check the common
         If bType = 6 Or bType = 7 Then
            sSql = "SELECT COCGSEXPACCT" & Trim(str(bType)) & "," _
                   & "COINVEXPACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         Else
            sSql = "SELECT COCGSMATACCT" & Trim(str(bType)) & "," _
                   & "COINVMATACCT" & Trim(str(bType)) & " " _
                   & "FROM ComnTable WHERE COREF=1"
         End If
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
         If bSqlRows Then
            With RdoAct
               If sDebitAcct = "" Then sDebitAcct = "" & Trim(.Fields(0))
               If sCreditAcct = "" Then sCreditAcct = "" & Trim(.Fields(1))
               .Cancel
            End With
         End If
      End If
   End If
   'After this excercise, we'll give up if none are found
   Set RdoAct = Nothing
   Exit Sub
   
DiaErr1:
   'Just bail for GetServerDateTime. May not have anything set
   'CurrError.Number = Err
   'CurrError.Description = Err.Description
   'DoModuleErrors Me
   On Error GoTo 0
End Sub

'Public Function GetServerDateTime() As Date
'    Dim prdoTme As ADODB.Recordset
'    sSql = "SELECT GETDATE()"
'    gblnSqlRows = clsADOCon.GetDataSet(sSql,prdoTme)
'    If gblnSqlRows Then
'        With prdoTme
'            GetServerDateTime = .Fields(0)
'        End With
'    End If
'    Set prdoTme = Nothing
'End Function
'

Public Function GetServerDateTimeToMinute() As Date
   Dim dt As Date
   dt = GetServerDateTime
   dt = DateAdd("s", -DatePart("s", dt), dt)
   GetServerDateTimeToMinute = dt
End Function

Public Function LoggedInToThisJob(sPartNo As String, nRunNo As Integer, nOpNo As Integer) As Boolean
   Dim i As Integer
   For i = 0 To UBound(mempCurrentEmployee.jobCurMO)
      Debug.Print "check " & mempCurrentEmployee.jobCurMO(i).strPart _
         & " MO " & mempCurrentEmployee.jobCurMO(i).lngRun _
         & " OP " & mempCurrentEmployee.jobCurMO(i).intOp
      If StrComp(mempCurrentEmployee.jobCurMO(i).strPart, sPartNo, vbTextCompare) = 0 _
                   And mempCurrentEmployee.jobCurMO(i).lngRun = nRunNo _
                   And mempCurrentEmployee.jobCurMO(i).intOp = nOpNo Then
         LoggedInToThisJob = True
         MsgBox "Already logged in to this job", vbInformation
         Exit Function
      End If
   Next
   LoggedInToThisJob = False
End Function

Public Sub LogInToJob(sPartNo As String, nRunNo As Integer, nOpNo As Integer, Optional bSetupTime As Boolean = False)
   Dim strSuRun As String
   
   gstrCurRoutine = "LogInToJob"
   On Error GoTo DiaErr1
   
   SystemAlert "Logging " & mempCurrentEmployee.strFirstName & " " _
      & mempCurrentEmployee.strLastName & " On To Jobs", , True
   
   ' close indirect time charge, if any before assigning
   ' the job.
   If mempCurrentEmployee.jobCurMO(0).lngRun = 0 Then
      CloseIndirectTC mempCurrentEmployee
   End If
   
   strSuRun = "R"
   If (bSetupTime = True) Then strSuRun = "S"
   
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   ' Add partial time charges to database...
   sSql = "INSERT INTO IstcTable " & vbCrLf _
          & "(ISEMPLOYEE,ISMO,ISRUN,ISOP,ISMOSTART,ISSHOP,ISWCNT,ISSURUN) " & vbCrLf _
          & "VALUES (" & mempCurrentEmployee.intNumber & ",'" _
          & Compress(sPartNo) & "'," _
          & nRunNo & "," _
          & nOpNo & ",'" _
          & GetServerDateTimeToMinute() & "','" & mempCurrentEmployee.strCurShop _
          & "','" & mempCurrentEmployee.strCurWC & " ','" & strSuRun & "')"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
   Else
      clsADOCon.RollbackTrans
      MsgBox Err
   End If
   Exit Sub
DiaErr1:
   gstrCurRoutine = "LogInToJob"
   DisplayError
   
End Sub


''0 = SysId
''1 = UserId
''2 = PartRef
''3 = ADate
''4 = Remaining Qty
''5 = Cost
''6 = Qty Selected
''Pass the finished array
'Public Sub UpdateLotArray()
'    Dim i As Integer
'    Dim a As Integer
'
'
'    On Error GoTo DiaErr1
'DiaErr1:
'
'End Sub


Public Function GetCustomReport(StdReport As String) As String
   Dim rdoCst As ADODB.Recordset
   Dim sNewReport As String
   'Strip the extension
   If Len(StdReport) > 4 Then
      If LCase$(Right$(StdReport, 4)) = ".rpt" Then _
                StdReport = Left$(StdReport, Len(StdReport) - 4)
   End If
   sNewReport = Trim(LCase$(StdReport))
   On Error GoTo modErr1
'   sSql = "SELECT REPORT_INDEX,REPORT_REF,REPORT_SECTION,REPORT_CUSTOMREPORT " _
'          & "FROM CustomReports WHERE REPORT_SECTION LIKE 'POM%' " _
'          & "AND REPORT_REF='" & Compress(sNewReport) & "'"
   sSql = "SELECT REPORT_CUSTOMREPORT " _
          & "FROM CustomReports WHERE REPORT_REF='" & Compress(sNewReport) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then sNewReport = Trim(rdoCst!REPORT_CUSTOMREPORT)
   If Trim(sNewReport) = "" Then sNewReport = LCase$(StdReport)
   GetCustomReport = sNewReport & ".rpt"
   Exit Function
   
modErr1:
   'it failed
   GetCustomReport = LCase$(StdReport) & ".rpt"
   
End Function


Sub CloseFiles()
   
End Sub


Public Function GetSetupTimeEnabled()
   Dim rdo As ADODB.Recordset
   Dim companyAccount As String
   
   sSql = "select ISNULL(COSETUPTCPOM, 0) as COSETUPTCPOM from ComnTable"
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If gblnSqlRows Then
      GetSetupTimeEnabled = rdo!COSETUPTCPOM
   Else
      GetSetupTimeEnabled = 0
   End If
   rdo.Close
End Function

Public Function GetAutoLogputOnPunchout()
   Dim rdo As ADODB.Recordset
   Dim companyAccount As String
   
   sSql = "select ISNULL(COLOGOUTONPOUT, 0) as COLOGOUTONPOUT from ComnTable"
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If gblnSqlRows Then
      GetAutoLogputOnPunchout = rdo!COLOGOUTONPOUT
   Else
      GetAutoLogputOnPunchout = 0
   End If
   rdo.Close
End Function


'Public Function InTestMode() As Boolean
'   If Dir(App.Path & "\UseTestDatabase.txt") <> "" Then
'      InTestMode = True
'   End If
'End Function

Public Sub LoadComboBox(Cntrl As Control, Optional ColumnNumber As Integer)
   'fill a combo box with the results of a query in sSql
   'For historic (stupid) reasons, the column is -1 based.
   'if you want the first column, pass -1
   'if no column is specified, the second column is used
   
   Dim ComboLoad As ADODB.Recordset
   Cntrl.Clear
'Debug.Print "LoadComboBox " & Cntrl.Name & " Clear count = " & Cntrl.ListCount
   If sSql = "" Then Exit Sub
   ColumnNumber = ColumnNumber + 1
   Set ComboLoad = clsADOCon.GetRecordSet(sSql, ES_STATIC)
'   Set ComboLoad = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
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
      Cntrl.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   sSql = ""
   Set ComboLoad = Nothing
   
End Sub

Public Sub AddComboStr(lhWnd As Long, sString As String)
   SendMessageStr lhWnd, CB_ADDSTRING, 0&, ByVal "" & Trim(sString)

End Sub

