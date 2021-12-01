Attribute VB_Name = "EsiTime"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Customer permissions 6/29/03
'10/7/04 added GetPoDataFormat
'1/20/05 added New Services PO
'3/31/05 Removed Jet excepting AWI
'8/8/05 Checked KeySet Clearing
'10/31/05 Added Cur.CurrentGroup to OpenFavorite. Opens appropriate tab
'         when called from Recent/Favorites and closed.
'12/22/05 Removed Vendor RFQ and references
'12/26/05 Removed unused procedures
'1/12/06 Completed renaming dialogs to be consistent with Fina
'5/3/06  Converted MRP columns to Dec(12,4)
'5/16/06 Delete Triggers on RunsTable/PohdTable
'6/5/06 BuildKeys
'6/26/06 Removed Threed32.ocx
'8/8/06 SSTab32.OCX Free
'9/7/06 AWI Custom MO - Removed the last JET references
'1/10/07 Added GetThisVendor for reports 7.2.1
Option Explicit
Public bGoodSoMo As Byte
Public bPOCaption As Byte
Public bFoundPart As Byte
Public iAutoIncr As Integer

Public sCurrDate As String
Public sCurrEmployee As String
Public sCurrForm As String
Public sPassedRout As String
Public sPassedMo As String
Public sPassedPart As String
Public sSelected As String

Public sFavorites(13) As String
Public sRecent(10) As String
Public sSession(50) As String

Public vTimeFormat As Variant

'Column updates
Private RdoCol As ADODB.Recordset
Private AdoError As ADODB.Error
Public gblnSqlRows As Boolean


'Private ER As rdoError

'Type tUser
'    Adduser As Integer
'    Level   As Integer
'    Group1  As Integer
'    Group2  As Integer
'    Group3  As Integer
'    Group4  As Integer
'    Group5  As Integer
'    Group6  As Integer
'    Group7  As Integer
'    Group8  As Integer
'End Type
'Public User As tUser
'
'8/22/05

Public Function GetUserLotID(UserLot As String) As Byte
   Dim RdoLot As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT DISTINCT LOTUSERLOTID FROM LohdTable WHERE " _
          & "LOTUSERLOTID='" & UserLot & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLot, ES_FORWARD)
   If bSqlRows Then GetUserLotID = 1 _
                                   Else GetUserLotID = 0
   ClearResultSet RdoLot
   If GetUserLotID = 1 Then MsgBox "That User Lot ID Is In Use.", _
                     vbInformation, "Revise A User Lot Number"
   Set RdoLot = Nothing
   
End Function

Public Function FindToolList(ToolNumber As String, ToolDesc As String, Optional _
                             DontShow As Byte) As String
   Dim RdoTlst As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "Qry_GetToolList '" & Compress(ToolNumber) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTlst, ES_FORWARD)
   If bSqlRows Then
      With RdoTlst
         FindToolList = "" & Trim(!TOOLLIST_NUM)
         If DontShow = 0 Then MDISect.ActiveForm.lblLst = "" & Trim(!TOOLLIST_DESC)
         ClearResultSet RdoTlst
      End With
   Else
      On Error Resume Next
      FindToolList = ""
      If DontShow = 0 Then MDISect.ActiveForm.lblLst = "*** Tool List Wasn't Found ***"
   End If
   Set RdoTlst = Nothing
   Exit Function
   
ModErr1:
   sProcName = "findtoollist"
   FindToolList = ""
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


Public Function GetRoutCenter(CenterRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "Qry_GetRoutCenter '" & CenterRef & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutCenter = "" & Trim(!WCNNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutCenter = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
ModErr1:
   sProcName = "getroutcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
   
End Function

'4/7/04

Public Function GetRoutShop(ShopRef As String) As String
   Dim RdoShp As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "Qry_GetShop '" & ShopRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         GetRoutShop = "" & Trim(!SHPNUM)
         ClearResultSet RdoShp
      End With
   Else
      GetRoutShop = ""
   End If
   Set RdoShp = Nothing
   Exit Function
   
ModErr1:
   sProcName = "getroutshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Function


'11/14/02 to retrieve the next PKRECORD (index piece)

Public Function GetNextPickRecord(sMoPartRef As String, lRunno As Long) As Integer
   Dim RdoRec As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT MAX(PKRECORD) FROM MopkTable WHERE " _
          & "PKMOPART='" & sMoPartRef & "' AND " _
          & "PKMORUN=" & lRunno & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRec, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoRec.Fields(0)) Then
         GetNextPickRecord = RdoRec.Fields(0) + 1
      Else
         GetNextPickRecord = 1
      End If
   Else
      GetNextPickRecord = 1
   End If
   Set RdoRec = Nothing
   Exit Function
   
ModErr1:
Resume modErr2:
modErr2:
   GetNextPickRecord = 1
   On Error GoTo 0
   Set RdoRec = Nothing
End Function


Public Sub GetLastMrp()
   Dim RdoShp As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT DISTINCT MRP_ROW,MRP_PARTDATERQD,MRP_USER " _
          & "FROM MrplTable WHERE MRP_ROW=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp, ES_FORWARD)
   If bSqlRows Then
      With RdoShp
         MDISect.ActiveForm.lblMrp = Format(!MRP_PARTDATERQD, "mm/dd/yy hh:mm AM/PM")
         MDISect.ActiveForm.lblUsr = "" & Trim(!MRP_USER)
         MDISect.ActiveForm.lblMrp.ForeColor = Es_TextForeColor
         ClearResultSet RdoShp
      End With
   Else
      MDISect.ActiveForm.lblMrp = "No Current Mrp"
      MDISect.ActiveForm.lblMrp.ForeColor = ES_RED
      MDISect.ActiveForm.lblUsr = ""
   End If
   Set RdoShp = Nothing
   Exit Sub
   
ModErr1:
   On Error GoTo 0
   
End Sub

Function FormatScheduleTime(Optional cHours As Currency) As Variant
   If cHours = 0 Then cHours = 8
   Select Case cHours
      Case Is < 8.5
         FormatScheduleTime = "mm/dd/yy 14:30"
      Case 8.5 To 16
         FormatScheduleTime = "mm/dd/yy 21:30"
      Case Is > 16
         FormatScheduleTime = "mm/dd/yy 23:59"
   End Select
   vTimeFormat = FormatScheduleTime
   
End Function


'11/21/06 Added .ListCount

Public Sub FillRoutings()
   Dim RdoRtg As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "Qry_FillRoutings "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRtg, ES_FORWARD)
   If bSqlRows Then
      With RdoRtg
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbRte.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRtg
      End With
   End If
   If MDISect.ActiveForm.cmbRte.ListCount > 0 Then _
      MDISect.ActiveForm.cmbRte.Text = MDISect.ActiveForm.cmbRte.List(0)
   Set RdoRtg = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "fillroutings"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

'8/9/05 Close open sets

Public Sub FormUnload(Optional bDontShowForm As Byte)
   Dim iList As Integer
   Dim iResultSets As Integer
   On Error Resume Next
   MDISect.lblBotPanel = MDISect.Caption
   'If Forms.count < 3 Then
   '   iResultSets = RdoCon.rdoResultsets.count
   '   For iList = iResultSets - 1 To 0 Step -1
   '      RdoCon.rdoResultsets(iList).Close
   '   Next
   'End If
   If bDontShowForm = 0 Then
      Select Case cUR.CurrentGroup
         Case "Time"
            SectionTimeEntry.Show
      End Select
      Erase bActiveTab
      cUR.CurrentGroup = ""
   End If
   
End Sub

'Find a favorite from the list

Public Sub OpenFavorite(sSelected As String)
   CloseForms
   If LTrim$(sSelected) = "" Then Exit Sub
   MouseCursor 13
   On Error GoTo OpenFavErr1
       Select Case sSelected
           Case "Time Type Codes"
               cUR.CurrentGroup = "Time"
               diaHcode.Show
           Case "Delete A Daily Time Charge"
               cUR.CurrentGroup = "Time"
               diaHdlch.Show
           Case "Operation Completions"
               cUR.CurrentGroup = "Time"
               diaHmops.Show
           Case "Enter/Revise Daily Time Charges"
               cUR.CurrentGroup = "Time"
               diaHrtme.Show
           Case "Employees By Name"
               cUR.CurrentGroup = "Time"
               diaPhu01.Show
           Case "Employees By Number"
               cUR.CurrentGroup = "Time"
               diaPhu02.Show
           Case "Daily Employee Time Charges"
               cUR.CurrentGroup = "Time"
               diaPhu03.Show
           Case "Weekly Time Charges (Report)"
               cUR.CurrentGroup = "Time"
               diaPhu05.Show
           Case "Time Type Codes Report"
               cUR.CurrentGroup = "Time"
               diaPhu15.Show
           Case "Daily Time Charges (Report)"
               cUR.CurrentGroup = "Time"
               diaPhu17.Show
           Case "Time Type Codes Report"
               cUR.CurrentGroup = "Time"
               diaPhu15.Show
           Case "Time Type Codes Report"
               cUR.CurrentGroup = "Time"
               diaPhu15.Show
           Case "Shift Codes"
               cUR.CurrentGroup = "Time"
               diaSfcode.Show
           Case "Employee Shift Code"
               cUR.CurrentGroup = "Time"
               diaSfEmp.Show
           Case "Apply Shift Code to  Daily Time Charges"
               cUR.CurrentGroup = "Time"
               diaSfHrtime.Show
           Case Else
               MouseCursor 0
       End Select
   MouseCursor 0
   On Error GoTo 0
   Exit Sub
   
OpenFavErr1:
   Resume OpenFavErr2
OpenFavErr2:
   MouseCursor 0
   MsgBox "ActiveX Error. Can't Load Form..", 48, "System    "
   On Error GoTo 0
   
End Sub

'11/21/06 Changed Name from GetDefaults

Public Sub GetRoutingIncrementDefault()
   Dim RdoDef As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT RTEINCREMENT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDef)
   If bSqlRows Then
      iAutoIncr = RdoDef!RTEINCREMENT
   Else
      iAutoIncr = 10
   End If
   If iAutoIncr = 0 Then iAutoIncr = 10
   Set RdoDef = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "getdefaults"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
End Sub






Sub Main()
   Dim sApptitle As String
   If App.PrevInstance Then
      On Error Resume Next
      sApptitle = App.Title
      App.Title = "E1ePr"
      SysMsgBox.Width = 3800
      SysMsgBox.msg.Width = 3200
      SysMsgBox.tmr1.Enabled = True
      SysMsgBox.msg = sApptitle & " Is Already Open."
      SysMsgBox.Show
      Sleep 5000
      AppActivate sApptitle
   End
   Exit Sub
End If
On Error Resume Next
' Set the Module name before loading the form
sProgName = "Time Management"
MainLoad "prod"
GetFavorites "EsiTime"
' MM 9/10/2009
'sProgName = "Time Management"
MDISect.Show

End Sub


'Replaced (Mostly) by GetCurrentPart

Public Sub FindPart(sGetPart As String, Optional NoMessage As Byte)
   Dim RdoPrt As ADODB.Recordset
   sGetPart = Compress(sGetPart)
   On Error GoTo ModErr1
   If Len(sGetPart) > 0 Then
      sSql = "Qry_GetINVCfindPart '" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      On Error Resume Next
      If bSqlRows Then
         With RdoPrt
            MDISect.ActiveForm.cmbPrt = "" & Trim(!PARTNUM)
            MDISect.ActiveForm.lblDsc = "" & !PADESC
            MDISect.ActiveForm.lblTyp = Format(0 + !PALEVEL, "0")
            MDISect.ActiveForm.lblUom = "" & Trim(!PAUNITS)
         End With
         bFoundPart = 1
      Else
         If NoMessage = 0 Then
            MsgBox "Part Wasn't Found.", 48, MDISect.ActiveForm.Caption
            MDISect.ActiveForm.cmbPrt = ""
         End If
         MDISect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
         MDISect.ActiveForm.lblTyp = ""
         bFoundPart = 0
      End If
   Else
      On Error Resume Next
      If NoMessage = 0 Then
         MDISect.ActiveForm.cmbPrt = "NONE"
         MDISect.ActiveForm.lblDsc = "*** Part Number Wasn't Found ***"
      End If
      bFoundPart = 0
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "findpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bFoundPart = 0
   DoModuleErrors MDISect.ActiveForm
   
End Sub



'Changed to controls 11/6/04

Public Function FindVendor(ContrlCombo As Control, ControlLabel As Control) As Byte
   Dim RdoVed As ADODB.Recordset
   Dim sVendRef As String
   sVendRef = Compress(ContrlCombo)
   If Len(sVendRef) = 0 Then Exit Function
   On Error GoTo ModErr1
   sSql = "Qry_GetVendorBasics '" & sVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVed)
   If bSqlRows Then
      On Error Resume Next
      With RdoVed
         ContrlCombo = "" & Trim(!VENICKNAME)
         ControlLabel = "" & Trim(!VEBNAME)
         FindVendor = True
         ClearResultSet RdoVed
      End With
   Else
      On Error Resume Next
      ContrlCombo = ""
      ControlLabel = "No Valid Vendor Selected."
      FindVendor = False
   End If
   Set RdoVed = Nothing
   Exit Function
   
ModErr1:
   sProcName = "findvendor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   FindVendor = False
   DoModuleErrors MDISect.ActiveForm
   
End Function

Public Sub FillRuns(frm As Form, sSearchString As String, Optional sComboName As String)
   Dim RdoFrn As ADODB.Recordset
   If sComboName = "" Then sComboName = "cmbPrt"
   On Error GoTo ModErr1
   If sSearchString = "<> 'CA'" Then
      sSql = "Qry_RunsNotCanceled"
   Else
      sSql = "Qry_RunsNotLikeC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFrn, ES_FORWARD)
   If bSqlRows Then
      With RdoFrn
         If sComboName = "cmbPrt" Then
            If Trim(frm.cmbPrt) = "" Then frm.cmbPrt = "" & Trim(!PARTNUM)
            Do Until .EOF
               AddComboStr frm.cmbPrt.hwnd, "" & Trim(!PARTNUM)
               .MoveNext
            Loop
         Else
            Do Until .EOF
               AddComboStr frm.cmbMon.hwnd, "" & Trim(!PARTNUM)
               .MoveNext
            Loop
         End If
         ClearResultSet RdoFrn
      End With
   End If
   On Error Resume Next
   Set RdoFrn = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "fillruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors frm
   
End Sub

'Use To Add Columns to a table where necessary.
'Will only update if the Column doesn't exist or if
'SQL Server isn't open. The later won't make any difference
'anyway.
'Create tables, indexes,columns, etc here

Public Sub UpDateTables()
   If MDISect.bUnloading = 1 Then Exit Sub
   Dim RdoTest As ADODB.Recordset
   
   MouseCursor 13
   SaveSetting "Esi2000", "AppTitle", "time", "ESI Time"
   SysOpen.Show
   SysOpen.prg1.Visible = True
   SysOpen.pnl = "Configuration Settings."
   SysOpen.pnl.Refresh
   
   On Error Resume Next
   SysOpen.prg1.Value = 20
   'moved to OldUpdate  1/3/03, 2/5/03, 5/9/05, 10/6/06
   
   '5/15/03 Patch for On Dock
   '*Leave in
   Err.Clear
   sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable WHERE VEREF='NONE'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTest, ES_KEYSET)
   If Not bSqlRows Then
      With RdoTest
         .AddNew
         !VEREF = "NONE"
         !VENICKNAME = "NONE"
         !VEBNAME = "No Vendor Selected"
         .Update
      End With
   End If
   CheckTriggers
   Sleep 500
   '6/5/06
   Err.Clear
   BuildKeys
   
   '1/8/07 Email to Vendor AR 7.2.0
   Err.Clear
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT VEAREMAIL FROM VndrTable WHERE VEAREMAIL='fubar'"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      
      sSql = "ALTER TABLE VndrTable ADD VEAREMAIL VARCHAR(60) NULL DEFAULT('')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "UPDATE VndrTable SET VEAREMAIL='' WHERE VEAREMAIL IS NULL"
         clsADOCon.ExecuteSQL sSql
      End If
      
   End If
   Err.Clear
   clsADOCon.ADOErrNum = 0
   
   '3/20/07 7.3.0 Add Actual Routing information
   sSql = "SELECT RUNRTNUM FROM RunsTable WHERE RUNRTNUM='FOOBAR'"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then
      Err.Clear
      sSql = "ALTER TABLE RunsTable ADD " _
             & "RUNRTNUM CHAR(30) NULL DEFAULT('')," _
             & "RUNRTDESC CHAR(30) NULL DEFAULT('')," _
             & "RUNRTBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPBY CHAR(20) NULL DEFAULT('')," _
             & "RUNRTAPPDATE CHAR(8) NULL DEFAULT('')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         sSql = "UPDATE RunsTable SET RUNRTNUM=''," _
                & "RUNRTDESC='',RUNRTBY=''," _
                & "RUNRTAPPBY='',RUNRTAPPDATE='' " _
                & "WHERE RUNRTNUM IS NULL"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
   Err.Clear
   SysOpen.prg1.Value = 80
   Sleep 500
   GoTo modErr2
   Exit Sub
   
ModErr1:
   Resume modErr2
modErr2:
   Set RdoTest = Nothing
   Err.Clear
   On Error GoTo 0
   SysOpen.Timer1.Enabled = True
   SysOpen.prg1.Value = 100
   SysOpen.Refresh
   Sleep 500
   
End Sub


Public Sub FindMoPart()
   Dim RdoPrt As ADODB.Recordset
   Dim sGetPart As String
   
   sGetPart = Compress(MDISect.ActiveForm.cmbMon)
   On Error GoTo ModErr1
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         On Error Resume Next
         With RdoPrt
            MDISect.ActiveForm.cmbMon = "" & Trim(!PARTNUM)
            MDISect.ActiveForm.lblMon = "" & !PADESC
         End With
      Else
         MDISect.ActiveForm.cmbMon = "NONE"
         MDISect.ActiveForm.lblMon = "*** Part Number Wasn't Found ***"
      End If
      Set RdoPrt = Nothing
   End If
   Exit Sub
   
ModErr1:
   sProcName = "findmopart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   bFoundPart = 0
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Public Function GetCenterCalendar(frm As Form, Optional sMonth As String) As Boolean
   'See if there are calendars
   'Syntax is bGoodCal = GetCenterCalendar(Me, Format$(SomeDate,"mm/dd/yy")
   On Error Resume Next
   If sMonth = "" Then
      sMonth = Format(ES_SYSDATE, "mmm") & "-" & Format(ES_SYSDATE, "yyyy")
   Else
      sMonth = Format(sMonth, "mmm") & "-" & Format(sMonth, "yyyy")
   End If
   sSql = "SELECT WCCREF FROM WcclTable WHERE WCCREF='" & sMonth & "'"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected = 0 Then
      GetCenterCalendar = False
      MsgBox "There Are No Work Center Calendars " & vbCr _
         & "Open For This Period " & sMonth & ".", vbInformation, frm.Caption
   Else
      GetCenterCalendar = True
   End If
   
End Function

'Currency because it rounds better and is faster
'Use local errors

Public Function GetCenterCalHours(sMonth As String, sShop As String, sCenter As String, iDay As Integer) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim cResources As Currency
   
   sMonth = Format(sMonth, "mmm") & "-" & Format(sMonth, "yyyy")
   sSql = "Qry_GetWorkCenterTimes '" & sMonth & "','" & sShop & "','" & sCenter & "'," _
          & iDay & " "
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetCenterCalHours = .Fields(0)
         Else
            GetCenterCalHours = 0
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetCenterCalHours = 0
   End If
   Set RdoTme = Nothing
   
   
End Function

'no calendar...try the workcenters
'6/2/00
'bGoodTime = GetCenterHours("CENTER", 2)
'Local errors

Public Function GetCenterHours(sCenter As String, iDay As Integer) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim sWkDay As String
   Select Case iDay
      Case 2
         sWkDay = "MON"
      Case 3
         sWkDay = "TUE"
      Case 4
         sWkDay = "WED"
      Case 5
         sWkDay = "THU"
      Case 6
         sWkDay = "FRI"
      Case 7
         sWkDay = "SAT"
      Case Else
         sWkDay = "SUN"
   End Select
   
   sWkDay = "WCN" & sWkDay & "HR1+WCN" & sWkDay & "HR2+WCN" _
            & sWkDay & "HR3+WCN" & sWkDay & "HR4"
   sSql = "SELECT SUM(" & sWkDay & ") FROM WcntTable WHERE " _
          & "WCNREF='" & sCenter & "'"
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetCenterHours = .Fields(0)
         Else
            GetCenterHours = 0
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetCenterHours = 0
   End If
   Set RdoTme = Nothing
   
End Function

Public Function GetThisCalendar(sMonth As String, sShop As String, sCenter As String) As Boolean
   On Error Resume Next
   sMonth = Format(sMonth, "mmm") & "-" & Format(sMonth, "yyyy")
   sSql = "SELECT WCCREF,WCCSHOP,WCCCENTER FROM WcclTable WHERE " _
          & "WCCREF='" & sMonth & "' AND (WCCSHOP='" & sShop & "' " _
          & "AND WCCCENTER='" & sCenter & "')"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected = 0 Then
      GetThisCalendar = False
   Else
      GetThisCalendar = True
   End If
   
End Function




Public Function GetThisCoCalendar(dTdate As Date) As Boolean
   Dim RdoCal As ADODB.Recordset
   Dim sTMonth As String
   Dim sTYear As String
   Dim sTDay As Integer
   sTMonth = Format(dTdate, "mmm")
   sTYear = sTMonth & "-" & Format(dTdate, "yyyy")
   sTDay = Format(dTdate, "d")
   sSql = "SELECT COCREF,COCDAY FROM CoclTable WHERE " _
          & "COCREF='" & sTYear & "' AND COCDAY=" & sTDay & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   If bSqlRows Then
      GetThisCoCalendar = True
   Else
      GetThisCoCalendar = False
   End If
   Set RdoCal = Nothing
End Function

Public Function GetQMCalHours(dTdate As Date) As Currency
   Dim RdoTme As ADODB.Recordset
   Dim sTMonth As String
   Dim iTDay As Integer
   sTMonth = Format(dTdate, "mmm") & "-" & Format(dTdate, "yyyy")
   iTDay = Format(dTdate, "d")
   sSql = "Qry_GetCompanyCalendarTime '" & sTMonth & "'," & iTDay & " "
   Set RdoTme = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   If Not RdoTme.BOF And Not RdoTme.EOF Then
      With RdoTme
         If Not IsNull(.Fields(0)) Then
            GetQMCalHours = .Fields(0)
         Else
            GetQMCalHours = 1
         End If
         ClearResultSet RdoTme
      End With
   Else
      GetQMCalHours = 1
   End If
   Set RdoTme = Nothing
End Function


Public Sub FillAllRuns(Contrl As Control)
   Dim RdoRns As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable WHERE PARTREF=RUNREF ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRns, ES_FORWARD)
   If bSqlRows Then
      With RdoRns
         Do Until .EOF
            AddComboStr Contrl.hwnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRns
      End With
      If Contrl.ListCount > 0 Then Contrl = Contrl.List(0)
   End If
   Set RdoRns = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "fillallruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

'Scripting for table
'Note non-normalized portions for speed and efficiency
'3/16/01

Public Sub MRPScript()
   'Types:
   '   Incoming (+)
   '   1 = Beginning balance
   '      2 = PO Items
   '      3 = MO Completions
   '   Out Going (-)
   '      4 = SO Items
   '      5 = Picks
   '      6 = Bills (Used On and no PL yet)
   '
   
End Sub

Public Sub FillBuyers()
   Dim RdoByr As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "Qry_GetBuyerList"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         Do Until .EOF
            AddComboStr MDISect.ActiveForm.cmbByr.hwnd, "" & Trim(!BYNUMBER)
            .MoveNext
         Loop
         ClearResultSet RdoByr
      End With
   End If
   Set RdoByr = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "fillbuyers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm
   
End Sub

Public Sub GetCurrentBuyer(sBuyer As String, Optional HideLabel As Byte)
   Dim RdoByr As ADODB.Recordset
   
   On Error GoTo ModErr1
   sBuyer = UCase$(Compress(sBuyer))
   sSql = "SELECT BYNUMBER,BYLSTNAME,BYFSTNAME,BYMIDINIT FROM " _
          & "BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         MDISect.ActiveForm.cmbByr = "" & Trim(!BYNUMBER)
         If HideLabel = 0 Then
            MDISect.ActiveForm.lblByr = "" & Trim(!BYFSTNAME) _
                                        & " " & Trim(!BYMIDINIT) & " " & Trim(!BYLSTNAME)
         End If
         ClearResultSet RdoByr
      End With
   Else
      If Len(Trim(sBuyer)) > 0 Then
         MDISect.ActiveForm.lblByr = "*** Buyer Wasn't Found ***"
      Else
         MDISect.ActiveForm.lblByr = ""
      End If
   End If
   Set RdoByr = Nothing
   Exit Sub
   
ModErr1:
   sProcName = "getcurrentbuyer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Forms(1)
   
   
End Sub

Public Function GetMoOperation(MONUMBER As String, Runno As Long, iOpno As Integer) As Byte
   Dim RdoOpr As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,OPREF,OPRUN,OPNO " _
          & "FROM RunsTable,RnopTable WHERE (RUNREF=OPREF AND " _
          & "RUNNO=OPRUN) AND (RUNREF='" & MONUMBER & "' AND RUNNO=" & Runno _
          & " AND OPNO=" & iOpno & " AND RUNSTATUS<>'CA' AND " _
          & "RUNSTATUS<>'CL')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOpr, ES_FORWARD)
   If bSqlRows Then GetMoOperation = 1 Else GetMoOperation = 0
   Set RdoOpr = Nothing
   Exit Function
   
ModErr1:
   GetMoOperation = 0
   
End Function

'1/20/04 calculate the end of the month

Public Function GetMonthEnd(Optional vMonthEnd As Variant) As Variant
   Dim bMonth As Byte
   Dim bEnd As Byte
   Dim bYear As Integer
   Dim vTest As Variant
   
   On Error Resume Next
   'Trap to test empty vMonth
   vTest = Left(vMonthEnd, 1)
   If Err > 0 Then
      bMonth = Format(ES_SYSDATE, "m")
      bYear = Format(ES_SYSDATE, "yyyy")
   Else
      bMonth = Format(vMonthEnd, "m")
      bYear = Format(vMonthEnd, "yyyy")
   End If
   Select Case bMonth
      Case 1, 3, 5, 7, 8, 10, 12
         bEnd = 31
      Case 2
         bEnd = 28
      Case Else
         bEnd = 30
   End Select
   
   If bEnd = 28 Then
      If bYear = 2004 Or bYear = 2008 Or bYear = 2012 _
                 Or bYear = 2016 Or bYear = 2020 Or bYear = 2024 Then bEnd = 29
   End If
   vMonthEnd = Format(bMonth, "00") & "/" & bEnd & "/" & Right$(str$(bYear), 2)
   GetMonthEnd = vMonthEnd
   
End Function



'10/7/04

Public Function GetPODataFormat() As String
   Dim RdoFormat As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT PurchasedDataFormat FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFormat, ES_FORWARD)
   If bSqlRows Then
      With RdoFormat
         If Not IsNull(!PurchasedDataFormat) Then
            GetPODataFormat = "" & Trim(!PurchasedDataFormat)
         Else
            GetPODataFormat = ES_QuantityDataFormat
         End If
         ClearResultSet RdoFormat
      End With
   End If
   If GetPODataFormat = "" Then GetPODataFormat = ES_QuantityDataFormat
   Set RdoFormat = Nothing
   Exit Function
   
ModErr1:
   GetPODataFormat = ES_QuantityDataFormat
   
End Function

'04/01/05

Public Function GetCompanyCalendar() As Byte
   Dim RdoCal As ADODB.Recordset
   Dim sCalYear As String
   Dim sCalMonth As String
   
   On Error Resume Next
   sCalYear = Format$(Now, "yyyy")
   sCalMonth = Format$(Now, "mmm")
   sSql = "Qry_GetCompanyCalendar '" & sCalMonth & "-" & sCalYear & " '"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCal, ES_FORWARD)
   ClearResultSet RdoCal
   GetCompanyCalendar = bSqlRows
   Set RdoCal = Nothing
   
End Function

Public Function TestWeekEnd(CalMonth As String, CalDay As String, CalShop As String, _
                            Calcenter As String) As Integer
   Dim RdoWend As ADODB.Recordset
   sSql = "SELECT SUM(WCCSHH1+WCCSHH2+WCCSHH3+WCCSHH4) As TotHours " _
          & "FROM WcclTable WHERE (WCCREF='" & CalMonth & "' AND " _
          & "DATENAME(dw,WCCDATE)Like '" & CalDay & "%' AND " _
          & "WCCSHOP='" & CalShop & "' AND WCCCENTER='" & Calcenter & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoWend, ES_FORWARD)
   If bSqlRows Then
      With RdoWend
         If Not IsNull(!TotHours) Then
            TestWeekEnd = !TotHours
         Else
            TestWeekEnd = 0
         End If
      End With
   Else
      TestWeekEnd = 0
   End If
   If TestWeekEnd < 2 Then TestWeekEnd = 0 _
                                         Else TestWeekEnd = 1
   
   Set RdoWend = Nothing
End Function

Public Function GetQNMConversion() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT QueueMoveConversion FROM Preferences WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet, ES_FORWARD)
   If bSqlRows Then
      With RdoGet
         If Not IsNull(!QueueMoveConversion) Then
            GetQNMConversion = !QueueMoveConversion
         Else
            GetQNMConversion = 24
         End If
         .Cancel
      End With
      ClearResultSet RdoGet
   Else
      GetQNMConversion = 24
   End If
   Set RdoGet = Nothing
   Exit Function
   
ModErr1:
   GetQNMConversion = 24
   
End Function

Public Sub GetMRPCreateDates(DateCreated As String, DateThrough As String)
   Dim RdoDate As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT MRP_ROW,MRP_CREATEDATE,MRP_THROUGHDATE FROM " _
          & "MrpdTable WHERE MRP_ROW=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDate, ES_FORWARD)
   If bSqlRows Then
      With RdoDate
         If Not IsNull(!MRP_CREATEDATE) Then
            DateCreated = Format$(!MRP_CREATEDATE, "mm/dd/yy")
         Else
            DateCreated = Format$(ES_SYSDATE, "mm/dd/yy")
         End If
         If Not IsNull(!MRP_CREATEDATE) Then
            DateThrough = Format$(!MRP_THROUGHDATE, "mm/dd/yy")
         Else
            DateThrough = Format$(ES_SYSDATE, "mm/dd/yy")
         End If
         .Cancel
      End With
   Else
      DateCreated = "  "
      DateThrough = "  "
   End If
   Set RdoDate = Nothing
   Exit Sub
ModErr1:
   Err.Clear
   DateCreated = "  "
   DateThrough = "  "
   
End Sub

'6/5/06

Private Sub BuildKeys()
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX RunsTable.RunRef"
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum > 0 Then GoTo KeysErr1
   
   sSql = "DROP INDEX RunsTable.RunPart"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNNO INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RunsTable ADD Constraint PK_RunsTable_RUNREF PRIMARY KEY CLUSTERED (RUNREF,RUNNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX RnopTable.OpRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP INDEX RnopTable.OpPart"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUN INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnopTable ALTER COLUMN OPNO SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnopTable ADD Constraint PK_RnopTable_OPREF PRIMARY KEY CLUSTERED (OPREF,OPRUN,OPNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   'No cascading
   sSql = "ALTER TABLE RnopTable ADD CONSTRAINT FK_RnopTable_RunsTable FOREIGN KEY (OPREF,OPRUN) References RunsTable"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX ShopTable.ShpRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPREF CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE ShopTable ADD Constraint PK_ShopTable_OPREF PRIMARY KEY CLUSTERED (SHPREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX WcntTable.WcnRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP INDEX WcntTable.WcnShop"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNREF CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSHOP CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcntTable ADD Constraint PK_WcntTable_WCNREF PRIMARY KEY CLUSTERED (WCNREF,WCNSHOP) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   sSql = "ALTER TABLE WcntTable ADD CONSTRAINT FK_WcntTable_ShopTable FOREIGN KEY (WCNSHOP) References ShopTable ON UPDATE CASCADE"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX CoclTable.CocRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ADD Constraint PK_CoclTable_COCREF PRIMARY KEY CLUSTERED (COCREF,COCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX CoclTable.CocRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ALTER COLUMN COCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CoclTable ADD Constraint PK_CoclTable_COCREF PRIMARY KEY CLUSTERED (COCREF,COCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX CctmTable.CalRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CctmTable ALTER COLUMN CALREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE CctmTable ADD Constraint PK_CctmTable_CALREF PRIMARY KEY CLUSTERED (CALREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX WcclTable.WcnRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCREF CHAR(8) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHOP CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCCENTER CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCDAY SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE WcclTable ADD Constraint PK_WcclTable_COCREF PRIMARY KEY CLUSTERED (WCCREF,WCCSHOP,WCCCENTER,WCCDAY) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX RnalTable.AllRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP INDEX RnalTable.RaRun"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP INDEX RnalTable.RaSo"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RAREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RARUN INTEGER NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASO INTEGER NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASOITEM SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ALTER COLUMN RASOREV CHAR(2) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_ALLOCATIONREF PRIMARY KEY CLUSTERED (RAREF,RARUN,RASO,RASOITEM,RASOREV) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_RunsTable FOREIGN KEY (RAREF,RARUN) References RunsTable ON UPDATE CASCADE"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnalTable ADD Constraint PK_RnalTable_PartTable FOREIGN KEY (RAREF) References PartTable"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX RndlTable.DlsRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSNUM SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSRUNREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RndlTable ALTER COLUMN RUNDLSRUNNO INTEGER NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RndlTable ADD Constraint PK_RndlTable_ALLOCATIONREF PRIMARY KEY CLUSTERED (RUNDLSNUM,RUNDLSRUNREF,RUNDLSRUNNO) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE RndlTable ADD CONSTRAINT FK_RndlTable_PartTable FOREIGN KEY (RUNDLSRUNREF) References PartTable ON UPDATE CASCADE ON DELETE CASCADE"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX PohdTable.PohdRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PohdTable ALTER COLUMN PONUMBER INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PohdTable ADD Constraint PK_PohdTable_PONUMBER PRIMARY KEY CLUSTERED (PONUMBER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   sSql = "DELETE FROM PoitTable " & vbCr _
          & "FROM PoitTable LEFT JOIN PohdTable ON PoitTable.PINUMBER = PohdTable.PONUMBER " & vbCr _
          & "WHERE (PohdTable.PONUMBER Is Null)"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX PoitTable.PoitRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PINUMBER INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PIITEM SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREV CHAR(2) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE PoitTable ADD Constraint PK_PoitTable_PINUMBER PRIMARY KEY CLUSTERED (PINUMBER,PIITEM,PIREV) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE PoitTable ADD CONSTRAINT FK_PoitTable_PohdTable FOREIGN KEY (PINUMBER) References PohdTable ON UPDATE CASCADE ON DELETE CASCADE"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX MopkTable.PickRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "DROP INDEX MopkTable.PkRecord"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKMOPART CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKMORUN INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MopkTable ALTER COLUMN PKRECORD SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MopkTable ADD Constraint PK_MopkTable_MOPICK PRIMARY KEY CLUSTERED (PKMOPART,PKMORUN,PKRECORD) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MopkTable ADD CONSTRAINT FK_MopkTable_PartTable FOREIGN KEY (PKPARTREF) References PartTable"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MopkTable ADD CONSTRAINT FK_MopkTable_RunsTable FOREIGN KEY (PKMOPART,PKMORUN) References RunsTable"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RnopTable ADD CONSTRAINT FK_RnopTable_WcntTable FOREIGN KEY (OPCENTER,OPSHOP) References WcntTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RtopTable ADD CONSTRAINT FK_RtopTable_WcntTable FOREIGN KEY (OPCENTER,OPSHOP) References WcntTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   sSql = "DELETE FROM BuyrTable WHERE BYREF IS NULL"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX BuyrTable.BuyerRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyrTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyrTable ADD Constraint PK_BuyrTable_BUYERID PRIMARY KEY CLUSTERED (BYREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   
   Err.Clear
   sSql = "DELETE FROM BuycTable " & vbCr _
          & "FROM BuycTable LEFT JOIN BuyrTable ON BuycTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX BuycTable.BuyercRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuycTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuycTable ALTER COLUMN BYPRODCODE CHAR(6) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuycTable ADD Constraint PK_BuycTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYPRODCODE) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "DELETE FROM BuypTable " & vbCr _
          & "FROM BuypTable LEFT JOIN BuyrTable ON BuypTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX BuypTable.BuyerpRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuypTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuypTable ALTER COLUMN BYPARTNUMBER CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuypTable ADD Constraint PK_BuypTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYPARTNUMBER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "DELETE FROM BuyvTable " & vbCr _
          & "FROM BuyvTable LEFT JOIN BuyrTable ON BuyvTable.BYREF = BuyrTable.BYREF " & vbCr _
          & "WHERE (BuyrTable.BYREF Is Null)"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX BuyvTable.BuyervRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyvTable ALTER COLUMN BYREF CHAR(20) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyvTable ALTER COLUMN BYVENDOR CHAR(10) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyvTable ADD Constraint PK_BuyvTable_BUYERID PRIMARY KEY CLUSTERED (BYREF,BYVENDOR) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE BuycTable ADD CONSTRAINT FK_BuycTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuypTable ADD CONSTRAINT FK_BuypTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyvTable ADD CONSTRAINT FK_BuyvTable_BuyrTable FOREIGN KEY (BYREF) References BuyrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuycTable ADD CONSTRAINT FK_BuycTable_PcodTable FOREIGN KEY (BYPRODCODE) References PcodTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuypTable ADD CONSTRAINT FK_BuypTable_PartTable FOREIGN KEY (BYPARTNUMBER) References PartTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE BuyvTable ADD CONSTRAINT FK_BuyvTable_VndrTable FOREIGN KEY (BYVENDOR) References VndrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   sSql = "DELETE FROM RfvdTable WHERE RFVENDOR='' OR RFVENDOR IS NULL"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX RfvdTable.RfvdTable_Unique"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFNO CHAR(12) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFITNO INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RfvdTable ALTER COLUMN RFVENDOR CHAR(10) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE RfvdTable ADD Constraint PK_RfvdTable_RFREV PRIMARY KEY CLUSTERED (RFNO,RFITNO,RFVENDOR) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE RfvdTable ADD CONSTRAINT FK_RfvdTable_VndrTable FOREIGN KEY (RFVENDOR) References VndrTable ON UPDATE CASCADE ON DELETE CASCADE "
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   
   sSql = "DROP INDEX MrplTable.MrpRow"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrplTable ALTER COLUMN MRP_ROW INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrplTable ADD Constraint PK_MrplTable_MRPREF PRIMARY KEY CLUSTERED (MRP_ROW) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MrplTable ADD CONSTRAINT FK_MrplTable_PartTable FOREIGN KEY (MRP_PARTREF) References PartTable"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   sSql = "DROP INDEX MrppTable.MRPPartRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrppTable ALTER COLUMN MRP_PARTREF CHAR(30) NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrppTable ADD Constraint PK_MrppTable_MRPPARTREF PRIMARY KEY CLUSTERED (MRP_PARTREF) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   Err.Clear
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   sSql = "DROP INDEX MrpbTable.BillRef"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrpbTable ALTER COLUMN MRPBILL_ORDER SMALLINT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrpbTable ADD Constraint PK_MrpbTable_MRPPORDER PRIMARY KEY CLUSTERED (MRPBILL_ORDER) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   
   sSql = "ALTER TABLE MrpdTable ADD CONSTRAINT FK_MrpbTable_PartTable FOREIGN KEY (MRPBILL_PARTREF) References PartTable"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "DROP INDEX MrpdTable.MrpDateIdx"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrpdTable ALTER COLUMN MRP_ROW INT NOT NULL"
   clsADOCon.ExecuteSQL sSql
   
   sSql = "ALTER TABLE MrpdTable ADD Constraint PK_MrpdTable_MRPDATE PRIMARY KEY CLUSTERED (MRP_ROW) " _
          & "WITH FILLFACTOR=80 "
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then clsADOCon.CommitTrans Else clsADOCon.RollbackTrans
   Exit Sub
   
KeysErr1:
   On Error Resume Next
   clsADOCon.RollbackTrans
   
   
End Sub

'ShopTable
'WcntTable
'RunsTable
'See ConvertRunOps



Private Sub ConvertProductionColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start ShopTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next
   'SHPRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPRATE dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPRATE DEC(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPRATE'"
               clsADOCon.ExecuteSQL sSql
            Else
               GoTo EndProc
            End If
         End If
      End If
   End With
   'SHPOH
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOH dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  If Err > 0 Then
                     For Each AdoError In RdoCol.ActiveConnection.Errors
                        sconstraint = GetConstraint(AdoError.Description)
                        If sconstraint <> "" Then Exit For
                     Next AdoError
                  End If
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOH DEC(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOH'"
               clsADOCon.ExecuteSQL sSql
            End If
            
         End If
      End If
   End With
   'SHPOHTOTAL
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOHTOTAL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHTOTAL dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHTOTAL DEC(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOHTOTAL'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPOHRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPOHRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHRATE dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPOHRATE dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPOHRATE'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPSETUP
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSETUP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSETUP dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSETUP dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSETUP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPUNIT
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNIT dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNIT dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPUNIT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPSECONDS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSECONDS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSECONDS dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSECONDS dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSECONDS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPSUHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSUHRS dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPSUHRS dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPSUHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPUNITHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNITHRS dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPUNITHRS dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPUNITHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPQHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPQHRS dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPQHRS dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPQHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPMHRS
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPMHRS dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPMHRS dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPMHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SHPESTRATE
   sSql = "sp_columns @table_name=ShopTable,@column_name=SHPESTRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPESTRATE dec(5,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE ShopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE ShopTable ALTER COLUMN SHPESTRATE dec(5,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'ShopTable.SHPESTRATE'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End ShopTable
   Err.Clear
   'Start WcntTable
   'WCNOHFIXED
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNOHFIXED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHFIXED dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHFIXED dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNOHFIXED'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNOHPCT
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNOHPCT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHPCT dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNOHPCT dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNOHPCT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSTDRATE
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSTDRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSTDRATE dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSTDRATE dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSTDRATE'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNUNITHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNUNITHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNUNITHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNUNITHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNESTRATE
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNESTRATE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNESTRATE dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNESTRATE dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNESTRATE'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNQHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNQHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNQHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNQHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMHRS
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   ConvertProductionColumns2


EndProc:
   On Error Resume Next
   'Update Preferences
   sSql = "UPDATE Preferences SET ProdtoDecimalConvDate='" & Format(Now, "mm/dd/yy") & "' " _
          & "WHERE (ProdtoDecimalConvDate IS NULL AND PreRecord=1)"
   clsADOCon.ExecuteSQL sSql
   Set RdoCol = Nothing
End Sub


Private Sub ConvertProductionColumns2() 'Had to break this out because I was getting a compile error that the original procudure was too big
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start ShopTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next
   
   
   'WCNWEDHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATHR1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATHR2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATHR3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATHR4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATHR4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSUNMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSUNMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSUNMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSUNMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNMONMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNMONMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNMONMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNMONMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTUEMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTUEMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTUEMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTUEMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNWEDMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNWEDMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNWEDMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNWEDMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNTHUMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNTHUMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNTHUMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNTHUMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNFRIMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNFRIMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNFRIMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNFRIMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATMU1
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU1 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU1 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATMU2
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU2 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU2 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATMU3
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU3 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU3 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCNSATMU4
   sSql = "sp_columns @table_name=WcntTable,@column_name=WCNSATMU4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU4 dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcntTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcntTable ALTER COLUMN WCNSATMU4 DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcntTable.WCNSATMU4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End WcntTable
   'Start RunsTable
   'RUNMATL
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNMATL dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNMATL dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNMATL'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNLABOR
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNLABOR dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNLABOR dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNLABOR'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNSTDCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNSTDCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSTDCOST dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSTDCOST dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNSTDCOST'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNPKQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNPKQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPKQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPKQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNPKQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNEXP dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNEXP dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNEXP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNYIELD
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNYIELD"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNYIELD dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNYIELD DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNYIELD'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNBUDLAB
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDLAB"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDLAB dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDLAB DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDLAB'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNBUDMAT
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDMAT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDMAT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDMAT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDMAT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNBUDEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDEXP dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDEXP DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDEXP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNBUDOH
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDOH dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDOH DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDOH'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNBUDHRS
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNBUDHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDHRS dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNBUDHRS dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNBUDHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCHARGED
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCHARGED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHARGED dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHARGED DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCHARGED'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCOST dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCOST DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCOST'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNOHCOST
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNOHCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNOHCOST dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNOHCOST DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNOHCOST'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCMATL
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCMATL dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCMATL DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCMATL'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCEXP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCEXP dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCEXP DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCEXP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCHRS
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHRS dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCHRS dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNCLAB
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNCLAB"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCLAB dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNCLAB DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNCLAB'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNPARTIALQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNPARTIALQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPARTIALQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNPARTIALQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNPARTIALQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNSCRAP
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNSCRAP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSCRAP dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNSCRAP dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNSCRAP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNREWORK
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNREWORK"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREWORK dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREWORK dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNREWORK'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'RUNREMAININGQTY
   sSql = "sp_columns @table_name=RunsTable,@column_name=RUNREMAININGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREMAININGQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RunsTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RunsTable ALTER COLUMN RUNREMAININGQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RunsTable.RUNREMAININGQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End RunsTable
   ConvertRunOPs
   ConvertCalendarColumns

End Sub

Private Function GetConstraint(sDescription As String) As String
   Dim bByte As Byte
   GetConstraint = ""
   bByte = InStr(1, sDescription, "DF_")
   If bByte > 0 Then
      GetConstraint = Mid$(sDescription, bByte, Len(sDescription))
      bByte = InStr(1, GetConstraint, "'")
      If bByte > 0 Then GetConstraint = Left$(GetConstraint, bByte - 1)
   Else
      bByte = InStr(1, sDescription, "DEFZERO")
      If bByte > 0 Then
         sSql = "sp_unbindefault '" & RdoCol.Fields(2) & "." & RdoCol.Fields(3) & "'"
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   
   
End Function

Private Function CheckConvErrors() As Byte
   Dim iColCounter As Integer
   CheckConvErrors = 0
   For Each AdoError In RdoCol.ActiveConnection.Errors
      If Left(AdoError.Description, 5) = "22003" Then
         iColCounter = iColCounter + 1
         CheckConvErrors = 1
      End If
   Next AdoError
   
End Function


'RnopTable
'RnalTable
'MopkTable

Public Sub ConvertRunOPs()
   Dim bBadCol As Byte
   Dim sconstraint As String
   On Error Resume Next
   'Start RnopTable
   'OPSETUP
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSETUP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSETUP dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSETUP DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSETUP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPUNIT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNIT dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNIT DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPUNIT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPQHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPQHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPQHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPQHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPQHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPMHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPMHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPMHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPMHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPMHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPSVCUNIT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSVCUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSVCUNIT dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSVCUNIT dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSVCUNIT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPYIELD
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPYIELD"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPYIELD dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPYIELD DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPYIELD'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPSUHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSUHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSUHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSUHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSUHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPUNITHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPUNITHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNITHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPUNITHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPUNITHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPRUNHRS
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPRUNHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUNHRS dec(9,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPRUNHRS DEC(9,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPRUNHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPCHARGED
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCHARGED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCHARGED dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCHARGED DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCHARGED'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPSHMUL
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSHMUL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSHMUL dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSHMUL DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSHMUL'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPCOST
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCOST dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCOST DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCOST'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPOHCOST
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPOHCOST"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPOHCOST dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPOHCOST DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPOHCOST'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPCONCUR
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPCONCUR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCONCUR dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPCONCUR DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPCONCUR'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPACCEPT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPACCEPT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPACCEPT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPACCEPT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPACCEPT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPREJECT
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPREJECT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREJECT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREJECT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPREJECT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPSCRAP
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPSCRAP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSCRAP dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPSCRAP dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPSCRAP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'OPREWORK
   sSql = "sp_columns @table_name=RnopTable,@column_name=OPREWORK"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREWORK dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnopTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnopTable ALTER COLUMN OPREWORK dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnopTable.OPREWORK'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End RnopTable
   Err.Clear
   'Start RnspTable
   'SPLIT_SPLQTY
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLORIGQTY
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLORIGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLORIGQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLORIGQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLORIGQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLLABOR
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLLABOR dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLLABOR DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLLABOR'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLOH
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLOH dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLOH DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLOH'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLHRS
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLHRS dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLHRS dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'SPLIT_SPLEXP
   sSql = "sp_columns @table_name=RnspTable,@column_name=SPLIT_SPLEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLEXP dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnspTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnspTable ALTER COLUMN SPLIT_SPLEXP dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnspTable.SPLIT_SPLEXP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End RnspTable
   Err.Clear
   'Start RnalTable
   'RAQTY
   sSql = "sp_columns @table_name=RnalTable,@column_name=RAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE RnalTable ALTER COLUMN RAQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE RnalTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE RnalTable ALTER COLUMN RAQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'RnalTable.RAQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End RnalTable
   Err.Clear
   'Start MopkTable
   'PKPQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKPQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKPQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKPQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKPQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKAQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKAQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKAMT
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKAMT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAMT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKAMT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKAMT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKOHPCT
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKOHPCT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKOHPCT dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKOHPCT dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKOHPCT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKORIGQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKORIGQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKORIGQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKORIGQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKORIGQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKINADDERS
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKINADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKINADDERS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKINADDERS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKINADDERS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKBOMQTY
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKBOMQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKBOMQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKBOMQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKBOMQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKTOTMATL
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTMATL"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTMATL dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTMATL DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTMATL'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKTOTLABOR
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTLABOR"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTLABOR dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTLABOR DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTLABOR'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKTOTEXP
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTEXP"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTEXP dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTEXP DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTEXP'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKTOTOH
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTOH"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTOH dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTOH DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTOH'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PKTOTHRS
   sSql = "sp_columns @table_name=MopkTable,@column_name=PKTOTHRS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTHRS dec(6,3)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE MopkTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE MopkTable ALTER COLUMN PKTOTHRS dec(6,3) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'MopkTable.PKTOTHRS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End MopkTable
   
End Sub

'PohdTable
'PoitTable
'VndrTable

Private Sub ConvertPurchasingColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start PohdTable //Test the first one and bail if not Real (see Else)
   On Error Resume Next
   'PODISCOUNT
   sSql = "sp_columns @table_name=PohdTable,@column_name=PODISCOUNT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'See Else
               sSql = "ALTER TABLE PohdTable ALTER COLUMN PODISCOUNT dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PohdTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PohdTable ALTER COLUMN PODISCOUNT dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PohdTable.PODISCOUNT'"
               clsADOCon.ExecuteSQL sSql
            Else
               GoTo EndProc
            End If
         End If
      End If
   End With
   'End PohdTable
   Err.Clear
   'Start PoitTable
   'PIPQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIPQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIPQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIPQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIPQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIAQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIAQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAQTY dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIAQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIAMT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIAMT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAMT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIAMT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIAMT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIESTUNIT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIESTUNIT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIESTUNIT dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIESTUNIT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIESTUNIT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIADDERS
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIADDERS dec(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIADDERS DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIADDERS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PILOT
   sSql = "sp_columns @table_name=PoitTable,@column_name=PILOT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PILOT DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PILOT DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PILOT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIFRTADDERS
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIFRTADDERS"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIFRTADDERS dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIFRTADDERS dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIFRTADDERS'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIREJECTED
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIREJECTED"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREJECTED DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIREJECTED DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIREJECTED'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIWASTE
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIWASTE"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIWASTE dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIWASTE dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIWASTE'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIORIGSCHEDQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIORIGSCHEDQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIORIGSCHEDQTY DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIORIGSCHEDQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIORIGSCHEDQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIONDOCKQTYACC
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIONDOCKQTYACC"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYACC DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYACC DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIONDOCKQTYACC'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIONDOCKQTYREJ
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIONDOCKQTYREJ"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYREJ DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIONDOCKQTYREJ DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIONDOCKQTYREJ'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'PIODDELQTY
   sSql = "sp_columns @table_name=PoitTable,@column_name=PIODDELQTY"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIODDELQTY DEC(12,4)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE PoitTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE PoitTable ALTER COLUMN PIODDELQTY DEC(12,4) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'PoitTable.PIODDELQTY'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End PoitTable
   Err.Clear
   'Start VndrTable
   'VEDISCOUNT
   sSql = "sp_columns @table_name=VndrTable,@column_name=VEDISCOUNT"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE VndrTable ALTER COLUMN VEDISCOUNT DEC(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE VndrTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE VndrTable ALTER COLUMN VEDISCOUNT DEC(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'VndrTable.VEDISCOUNT'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   
   
EndProc:
   On Error Resume Next
   'Update Preferences
   sSql = "UPDATE Preferences SET PurctoDecimalConvDate='" & Format(Now, "mm/dd/yy") & "' " _
          & "WHERE (PurctoDecimalConvDate IS NULL AND PreRecord=1)"
   clsADOCon.ExecuteSQL sSql
   RdoCol.Close
   
End Sub

Private Sub ConvertCalendarColumns()
   Dim bBadCol As Byte
   Dim sconstraint As String
   'Start WcclTable
   'WCCSHH1
   On Error Resume Next
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then 'see Else
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH1'"
               clsADOCon.ExecuteSQL sSql
            Else
               Exit Sub
            End If
         End If
      End If
   End With
   'WCCSHH2
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHH3
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHH4
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHH4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHH4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHH4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHR1
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHR2
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHR3
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'WCCSHR4
   sSql = "sp_columns @table_name=WcclTable,@column_name=WCCSHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE WcclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE WcclTable ALTER COLUMN WCCSHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'WcclTable.WCCSHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End WcclTable
   Err.Clear
   'Start CctmTable
   'CALSUNHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSUNHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSUNHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSUNHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSUNHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSUNHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSUNHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALMONHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALMONHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALMONHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALMONHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALMONHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALMONHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALMONHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTUEHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTUEHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTUEHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTUEHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTUEHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTUEHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTUEHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALWEDHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALWEDHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALWEDHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALWEDHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALWEDHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALWEDHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALWEDHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTHUHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTHUHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTHUHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALTHUHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALTHUHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALTHUHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALTHUHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALFRIHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALFRIHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALFRIHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALFRIHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALFRIHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALFRIHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALFRIHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSATHR1
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSATHR2
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSATHR3
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'CALSATHR4
   sSql = "sp_columns @table_name=CctmTable,@column_name=CALSATHR4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CctmTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CctmTable ALTER COLUMN CALSATHR4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CctmTable.CALSATHR4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End CctmTable
   Err.Clear
   'Start CoclTable
   'COCSHT1
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT1"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT1 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT1 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT1'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'COCSHT2
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT2"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT2 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT2 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT2'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'COCSHT3
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT3"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT3 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT3 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT3'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'COCSHT4
   sSql = "sp_columns @table_name=CoclTable,@column_name=COCSHT4"
   Set RdoCol = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   With RdoCol
      If Not .BOF And Not .EOF Then
         Err.Clear
         If Not IsNull(.Fields(5)) Then
            If .Fields(5) = "real" Then
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT4 dec(7,2)"
               clsADOCon.ExecuteSQL sSql
               If Err > 0 Then
                  For Each AdoError In RdoCol.ActiveConnection.Errors
                     sconstraint = GetConstraint(AdoError.Description)
                     If sconstraint <> "" Then Exit For
                  Next AdoError
               End If
               Err.Clear
               If sconstraint <> "" Then
                  sSql = "ALTER TABLE CoclTable DROP " & sconstraint
                  clsADOCon.ExecuteSQL sSql
               End If
               sSql = "ALTER TABLE CoclTable ALTER COLUMN COCSHT4 dec(7,2) "
               clsADOCon.ExecuteSQL sSql
               
               bBadCol = CheckConvErrors()
               sSql = "sp_bindefault DEFZERO, 'CoclTable.COCSHT4'"
               clsADOCon.ExecuteSQL sSql
            End If
         End If
      End If
   End With
   'End CoclTable
   
   
End Sub

'Moved here 10/6/06 Leave in

Private Sub CheckTriggers()
   Err.Clear
   '5/16/06 Delete Trigger (Runs)
   sSql = "CREATE TRIGGER DT_RunsTable ON RunsTable" & vbCr _
          & "FOR  DELETE " & vbCr _
          & "  AS " & vbCr _
          & "  SAVE TRANSACTION SaveRows " & vbCr _
          & "  Rollback TRANSACTION"
   clsADOCon.ExecuteSQL sSql
   
   Err.Clear
   '5/16/06 Delete Trigger (PO's)
   sSql = "CREATE TRIGGER DT_PohdTable ON PohdTable" & vbCr _
          & "FOR  DELETE " & vbCr _
          & "  AS " & vbCr _
          & "  SAVE TRANSACTION SavePoRows " & vbCr _
          & "  Rollback TRANSACTION"
   clsADOCon.ExecuteSQL sSql
   
End Sub

'1/10/07 Added for reports

Public Sub GetThisVendor()
   Dim RdoRpt As ADODB.Recordset
   On Error GoTo ModErr1
   sSql = "SELECT VEBNAME FROM VndrTable WHERE VEREF='" _
          & Compress(MDISect.ActiveForm.cmbVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      MDISect.ActiveForm.lblVEName = "" & Trim(RdoRpt!VEBNAME)
      ClearResultSet RdoRpt
   Else
      MDISect.ActiveForm.lblVEName = "*** A Range Of Vendors Selected ***"
   End If
   Set RdoRpt = Nothing
   Exit Sub
ModErr1:
   On Error GoTo 0
   
End Sub


Public Function IsValidRunOp(PartNo As String, Runno As String, opNo As String, Optional IncludeCancelled = True, Optional IncludeClosed = True) As Boolean
   'return True if this run operation is valid
   
   'OK if both run # and op# are blank
   Dim nRunNo As Integer
   Dim nOpNo As Integer
   
   ' Remove the space and evalute
   Runno = Replace(Runno, Chr$(32), "")
   opNo = Replace(opNo, Chr$(32), "")
   nRunNo = CInt("0" & Runno)
   nOpNo = CInt("0" & opNo)
   
   IsValidRunOp = False
   
   If nRunNo = 0 And nOpNo = 0 Then
      IsValidRunOp = True
      Exit Function
   End If
   
   
   On Error GoTo DiaErr1
   Dim rdo As ADODB.Recordset
   sSql = "select *" _
          & " from RunsTable run" _
          & " join RnopTable op on op.OPREF = run.RUNREF" _
          & " and op.OPRUN = run.RUNNO" _
          & " where OPREF='" & Compress(PartNo) _
          & "' and OPRUN=" & CStr(nRunNo) _
          & " and OPNO=" & CStr(nOpNo)
          
    If Not IncludeCancelled Then sSql = sSql & " AND RUNSTATUS <> 'CA'"
    If Not IncludeClosed Then sSql = sSql & " AND RUNSTATUS<> 'CL'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      IsValidRunOp = True
   End If
   
   Exit Function
   
DiaErr1:
   sProcName = "IsValidRunOp"
   IsValidRunOp = False
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MDISect.ActiveForm

End Function


Public Sub FindShop(frm As Form)
   'Use local errors
   Dim sGetShop As String
   Dim RdoShp As ADODB.Recordset
   sGetShop = Compress(frm.cmbShp)
   If Len(sGetShop) > 0 Then
      sSql = "SELECT SHPREF,SHPNUM FROM ShopTable WHERE SHPREF='" & sGetShop & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoShp)
      If bSqlRows Then
         frm.cmbShp = "" & Trim(RdoShp!SHPNUM)
      Else
         MsgBox "Shop Wasn't Found.", 48, "Shops"
         frm.cmbShp = ""
      End If
      On Error Resume Next
      RdoShp.Cancel
      Set RdoShp = Nothing
   End If
   
End Sub
