VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTimeCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function IsValidTime(time As Variant)
   On Error Resume Next
   Dim n As Long
   n = DatePart("n", time)
   If Err Then
      IsValidTime = False
   Else
      IsValidTime = True
   End If
End Function

Public Function ComputeElapsedTime(startTime As Variant, endTime As Variant, ElapsedHours As Currency) As Boolean
   'if start time and stop time are valid, returns True
   'if either time is invalid, returns false
   'if starttime = endtime, return false
   'ElapsedHours = the elapsed time in hours
   
   ElapsedHours = 0
   'if invalid time, don't continue
   If Not IsValidTime(startTime) Or Not IsValidTime(endTime) Then
      Exit Function
   End If
   
   'if starttime = endtime, don't continue
   If DateDiff("n", startTime, endTime) = 0 Then
      Exit Function
   End If
   
   ElapsedHours = Round(DateDiff("n", startTime, endTime) / 60#, 3)
   ComputeElapsedTime = True

End Function

Public Function ConvertHours(cTime As Currency) As String
   'Converts real hours like 8.3 to time like 08:18
   'Note systax - Pass a number, returns a string or variant
   'sSomeChangedTime = ConvertHours(8.3)
   
   Dim min As Integer
   min = cTime * 60
   ConvertHours = DateAdd("n", min, "1/1/1900")
End Function

Public Function CreateTimeCharge(TimeCardID As String, empno As Long, _
   StartDateTime As Variant, EndDateTime As Variant, TimeCode As String, _
   SetupRunIndirect As String, IndirectAccount As String, _
   PartNumber As String, RunNumber As Long, OpNumber As Integer, _
   JournalRef As Integer, Source As String, _
   AcceptCount, RejectCount, ScrapCount, Comments As String, _
   Optional bMOComt As Boolean = False) As Boolean
   
   'IndirectAccount is required for indirect charges (SetupRunIndirec = "I")
   'return = True if successful
   
   Dim JournalID As String
   If Not GetOpenTimeJournalForThisDate(StartDateTime, JournalID) Then
      Exit Function
   End If
   
   If SetupRunIndirect = "I" And Trim(IndirectAccount) = "" Then
      MsgBox "Account must be specified for indirect charge"
      Exit Function
   End If
   
   If SetupRunIndirect = "I" And Trim(PartNumber) <> "" Then
      MsgBox "You cannot specify an MO for an indirect charge"
      Exit Function
   End If
   
   If SetupRunIndirect <> "I" And Trim(IndirectAccount) <> "" Then
      MsgBox "Account cannot be specified for a direct charge"
      Exit Function
   End If
   
   If SetupRunIndirect <> "I" And Trim(PartNumber) = "" Then
      MsgBox "MO must be specified for a direct charge"
      Exit Function
   End If
   
   Dim DebitAccount As String, Shop As String, WorkCenter As String
   Dim Rate As Currency, OverheadRate As Currency, OverheadFixed As Currency
   
   On Error GoTo whoops
   
   Dim Hours As Currency, hoursAsDateTime As String
   Dim StartDateTimeToMinute As Variant, EndDateTimeToMinute As Variant
   StartDateTimeToMinute = Format(StartDateTime, "mm/dd/yy hh:mm")
   EndDateTimeToMinute = Format(EndDateTime, "mm/dd/yy hh:mm")
   If Not ComputeElapsedTime(StartDateTimeToMinute, EndDateTimeToMinute, Hours) Then
      Exit Function
   End If
   
   hoursAsDateTime = ConvertHours(Hours)
   
   Dim sortKey As Long
   sortKey = CDate(Format(StartDateTimeToMinute, "hh:mm")) * 1000000       'sort order
   
   'if direct charge, look up rates and accounts
   If SetupRunIndirect = "I" Then
      GetEmployeeAccountAndRate empno, TimeCode, DebitAccount, Rate
      DebitAccount = IndirectAccount      'use account provided
   Else
      GetOpAccountAndRates PartNumber, RunNumber, OpNumber, empno, TimeCode, DebitAccount, Rate, _
         OverheadFixed, OverheadRate, Shop, WorkCenter
   
      ' MM OverHead rate is percentage...calculate the actual rate
      OverheadRate = (OverheadRate * Rate) / 100
   End If

   If DebitAccount = "" Then
      DebitAccount = GetDefTimeAccounts("Time")
   End If
   
   
   Dim startString As String, endString As String
   startString = GetTime(StartDateTime)
   endString = GetTime(EndDateTime)
   
   If (startString = "") Then
      startString = "12:00a"
   End If
   
   If (endString = "") Then
      endString = "12:00a"
   End If
   
   ' debugging dup time charges.  Make sure charge does not exist
   sSql = "select * from TcitTable where TCCARD='" & TimeCardID & "'" & vbCrLf _
      & "and TCPARTREF='" & PartNumber & "' and TCRUNNO=" & RunNumber & " and TCOPNO=" & OpNumber & vbCrLf _
      & "and TCSTARTTIME='" & StartDateTimeToMinute & "'"
   Dim rs As ADODB.Recordset
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   rs.Close
   If bSqlRows Then
'      MsgBox "Time charge already exists for MO " & PartNumber & "-" & RunNumber & " OP " & OpNumber & "." & vbCrLf _
'         & "Please inform an administrator"
      CreateTimeCharge = True
      Exit Function
   End If
   
   If Hours > 0 Then
      sSql = "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf _
         & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf _
         & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf _
         & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf _
         & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf _
         & "VALUES('" & TimeCardID & "'," & empno & ",'" & startString & "'," & vbCrLf _
         & "'" & endString & "'," & "'" & StartDateTimeToMinute & "', '" & EndDateTimeToMinute & "'," & vbCrLf _
         & Hours & ",'" & hoursAsDateTime & "'," & vbCrLf _
         & "'" & TimeCode & "'," & Rate & "," & OverheadRate & "," & vbCrLf _
         & "1,'" & Compress(DebitAccount) & "','" & Compress(DebitAccount) & "','" & Shop & "','" & WorkCenter & "'," _
         & "0,'" & SetupRunIndirect & "',0,'" & PartNumber & "'," _
         & RunNumber & "," & OpNumber & "," & sortKey & "," & OverheadFixed & "," & vbCrLf _
         & "'" & JournalID & "'," & JournalRef & ",'" & Source & "'," & vbCrLf _
         & "0, " & AcceptCount & ", " & RejectCount & ", " & ScrapCount & ",'" & Replace(Comments, "'", "''") & "')"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      sSql = "tried to rerun CreateTimeCharge SQL"
   End If
   
   If (bMOComt = True) Then
        AddMOOpComment PartNumber, RunNumber, Comments
   End If
   
   CreateTimeCharge = True
   Exit Function
   
whoops:
   sProcName = "CreateTimeCharge"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm

End Function

Private Function AddMOOpComment(PartNumber As String, RunNumber As Long, Comments As String)

    sSql = "Update RunsTable set RUNCOMMENTS = '" & Comments & "' " & vbCrLf _
        & " WHERE RUNREF = '" & Compress(PartNumber) & "' " & vbCrLf _
         & " AND RUNNO = " & RunNumber
   
   clsADOCon.ExecuteSql sSql ', rdExecDirect
            
End Function


Public Sub GetOpAccountAndRates(PartNo As String, ByVal Runno As Long, _
   opNo As Integer, empno As Long, TimeCode As String, Account As String, Rate As Currency, _
   FixedOhRate As Currency, PctOhRate As Currency, Shop As String, WorkCenter As String)
   
   'returned values:
   'Account (debit)
   'Rate
   'FixedOhRate
   'PctOhRate
   'Shop
   'WorkCenter
   
   Dim rdo As ADODB.Recordset
'
'   'get employee information
'   Dim empRate As Currency, empAccount As String
'   sSql = "select PREMPAYRATE, rtrim(PREMACCTS) as PREMACCTS from EmplTable where PREMNUMBER = " & EmpNo
'   gblnSqlRows = clsADOCon.GetDataSet(sSql,rdo)
'   If gblnSqlRows Then
'      empRate = rdo!PREMPAYRATE
'      empAccount = rdo!PREMACCTS
'   Else
'      empRate = 0
'      empAccount = ""
'   End If
'   rdo.Close
'

   Dim empRate As Currency, empAccount As String
   GetEmployeeAccountAndRate empno, TimeCode, empAccount, empRate
   
   'get default information
   Dim companyAccount As String
   sSql = "select rtrim(WIPLABACCT) as WIPLABACCT from ComnTable"
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If gblnSqlRows Then
      companyAccount = rdo!WIPLABACCT
   Else
      companyAccount = ""
   End If
   rdo.Close
   Set rdo = Nothing
   
   'get part/wc/shop information
   PartNo = Compress(PartNo)
   Dim wcAccount As String, wcFixedOh As Currency, wcPctOh As Currency, wcRate As Currency
   sSql = "select PAUSEACTUALCOST, WCNOHPCT, SHPOHRATE, WCNOHFIXED, SHPOHTOTAL," & vbCrLf _
          & "WCNSTDRATE, SHPRATE, rtrim(WCNACCT) as WCNACCT, rtrim(SHPACCT) as SHPACCT," & vbCrLf _
          & "rtrim(OPCENTER) as OPCENTER, rtrim(OPSHOP) as OPSHOP" & vbCrLf _
          & "from RnopTable r" & vbCrLf _
          & "join PartTable p on r.OPREF = p.PARTREF" & vbCrLf _
          & "join WcntTable w on r.OPCENTER = w.WCNREF" & vbCrLf _
          & "join ShopTable s on r.OPSHOP = s.SHPREF" & vbCrLf _
          & "where r.OPREF='" & PartNo & "'" & vbCrLf _
          & "and r.OPRUN=" & Runno & vbCrLf _
          & "and r.OPNO=" & opNo
   gblnSqlRows = clsADOCon.GetDataSet(sSql, rdo)
   If gblnSqlRows Then
      With rdo
         'get shop and workcenter
         Shop = !OPSHOP
         WorkCenter = !OPCENTER
         
         'determine account to use
         If !WCNACCT <> "" Then
            Account = !WCNACCT
         ElseIf !SHPACCT <> "" Then
            Account = !SHPACCT
         Else
            Account = companyAccount
         End If
         
         'determine rate to use
'         If !paUseActualCost <> 0 Then
'            Rate = empRate
'         ElseIf !WCNSTDRATE <> 0 Then
'            Rate = !WCNSTDRATE
'         Else
'            Rate = !SHPRATE
'         End If
         
         If empRate <> 0 Then
            Rate = empRate
         ElseIf !WCNSTDRATE <> 0 Then
            Rate = !WCNSTDRATE
         Else
            Rate = !SHPRATE
         End If
         
         
         
         'determine percent overhead to use
         If !WCNOHPCT <> 0 Then
            PctOhRate = !WCNOHPCT
         ElseIf !SHPOHRATE <> 0 Then
            PctOhRate = !SHPOHRATE
         End If
         
         'determine fixed overhead to use
         If !WCNOHFIXED <> 0 Then
            FixedOhRate = !WCNOHFIXED
         ElseIf !SHPOHTOTAL <> 0 Then
            FixedOhRate = !SHPOHTOTAL
         End If
         
         
      End With
      Set rdo = Nothing
      
'      'add shift differential to rate if not holiday, break or vacation
'      ' do this when getting basic employee rate only
'      If TimeCode <> "VC" And TimeCode <> "HL" And TimeCode <> "BR" Then
'         sSql = "select SFADDRT from sfempTable emp" & vbCrLf _
'            & "join sfcdTable cd on cd.SFCODE = emp.SFREFMON" & vbCrLf _
'            & "WHERE PREMNUMBER = " & CStr(empno)
'         Dim rs As ADODB.Recordset, differential As Currency
'         bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
'         If bSqlRows Then
'            differential = rs!SFADDRT
'            Rate = Rate + differential
'         End If
'         Set rs = Nothing
'      End If

         

   End If
   
'   'multiple by rate for overtime, doubletime
'   Dim multiplier As Currency
'   multiplier = Me.GetTimeCodeMultiplier(TimeCode)
'   If multiplier <> 0 Then
'      Rate = multiplier * Rate
'   End If
'
End Sub

Public Sub GetEmployeeAccountAndRate(empno As Long, TimeCode As String, empAccount As String, empRate As Currency)
   
   'get employee information
   Dim rdo As ADODB.Recordset
   sSql = "select PREMPAYRATE, rtrim(PREMACCTS) as PREMACCTS from EmplTable where PREMNUMBER = " & empno
   If clsADOCon.GetDataSet(sSql, rdo) Then
      empRate = rdo!PREMPAYRATE * Me.GetTimeCodeMultiplier(TimeCode)
      empAccount = rdo!PREMACCTS
   Else
      empRate = 0
      empAccount = ""
   End If
   Set rdo = Nothing
   
      'add shift differential to rate if not holiday, break or vacation
   If TimeCode <> "VC" And TimeCode <> "HL" And TimeCode <> "BR" Then
      sSql = "select SFADDRT from sfempTable emp" & vbCrLf _
         & "join sfcdTable cd on cd.SFCODE = emp.SFREFMON" & vbCrLf _
         & "WHERE PREMNUMBER = " & CStr(empno) & " and SFADDRT is not null"
      Dim rs As ADODB.Recordset, differential As Currency
      bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
      If bSqlRows Then
         differential = rs!SFADDRT
         empRate = empRate + differential
      End If
      Set rs = Nothing
   End If

End Sub

Public Function GetTimeCodeMultiplier(TimeCode As String) As Currency
   'return = 1 if no such timecode found or timecode has no multiplier
   
   Dim rdo As ADODB.Recordset
   sSql = "SELECT TYPEADDER FROM TmcdTable WHERE TYPECODE = '" & TimeCode & "'"
   If clsADOCon.GetDataSet(sSql, rdo) Then
      GetTimeCodeMultiplier = rdo!TYPEADDER
   End If
   
   If GetTimeCodeMultiplier = 0 Then
      GetTimeCodeMultiplier = 1
   End If
   Set rdo = Nothing
End Function

Public Function GetDefTimeAccounts(AccountType As String) As String
   'Time for Time Cards and Labor for distribution
   'AccountType = "Time" or "Labor"
   
   Dim RdoAcc As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT CODEFTIMEACCT,CODEFLABORACCT FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAcc, ES_FORWARD)
   If bSqlRows Then
      With RdoAcc
         If AccountType = "Time" Then
            GetDefTimeAccounts = "" & Trim(.Fields(0))
         Else
            GetDefTimeAccounts = "" & Trim(.Fields(1))
         End If
         ClearResultSet RdoAcc
      End With
   End If
   Set RdoAcc = Nothing
   Exit Function
   
modErr1:
   GetDefTimeAccounts = ""
   
End Function



Public Function GetTimeCardID(empno As Long, startDate As Date, Optional strSftBegTime As String = "")
   'find the existing timecard for employee and date
   'if none exists, create one
   
   Dim dateString As String
   
   If (Trim(strSftBegTime) <> "") Then
      dateString = Format(strSftBegTime, "mm/dd/yyyy")
   Else
      dateString = Format(startDate, "mm/dd/yyyy")
   End If
   sSql = "select TMCARD from TchdTable" & vbCrLf _
      & "where TMEMP = " & empno & " and TMDAY = '" & dateString & "'"
   Dim rdo As ADODB.Recordset
   If clsADOCon.GetDataSet(sSql, rdo) Then
      GetTimeCardID = rdo.Fields(0)
      Set rdo = Nothing
      Exit Function
   End If
   
   Set rdo = Nothing
   
   'timecard doesn't exist.  create it.
   Dim S As Single
   Dim l As Long
   Dim m As Long
   'Dim T As String
   'Dim timeCardNo As String
   Dim loginDate As String
   m = DateValue(Format(GetServerDateTime, "yyyy,mm,dd"))
   S = TimeValue(Format(GetServerDateTime, "hh:mm:ss"))
   l = S * 1000000
   GetTimeCardID = Format(m, "00000") & Format(l, "000000")
   sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDAY)" & vbCrLf _
      & "VALUES ('" & GetTimeCardID & "', " & empno & ", '" & dateString & "')"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
End Function

Public Function GetOpenTimeJournalForThisDate(dt As Variant, JournalID As String, _
   Optional DontDisplayErrorMsg As Boolean) As Boolean
   'place it in journalid
   'RETURN = True if successful
   
   Dim b As Boolean
   
   'ok if journal not required
   JournalID = GetOpenJournal("TJ", Format$(dt, "mm/dd/yy"))
   If Left(JournalID, 4) = "None" Then
      JournalID = ""
      b = True
   ElseIf JournalID = "" Then
      b = False
      If Not DontDisplayErrorMsg Then
         MsgBox "There is no open time journal for " & Format(dt, "mm/dd/yy") & ".", vbCritical
      End If
   Else
      b = True
   End If
   
'   If Not b Then
'      MsgBox "There Is No Open Time Journal For This Period.", vbExclamation
'   End If
   GetOpenTimeJournalForThisDate = b
End Function


Public Function GetTime(TimeEntry As Variant) As Variant
   'place entered string in hh:mma/p format
   'returns blank if invalid date
   'syntax  is txtTme = GetTime (txtTme)
   
   'get am/pm
   Dim i As Integer, ampm As String, lastDigit As Integer, ch As String, colon As Integer
   Dim digits As Integer, ts As String
   ampm = "a"
   For i = 1 To Len(TimeEntry)
      ch = Mid(TimeEntry, i, 1)
      If IsNumeric(ch) Then
         lastDigit = i
         digits = digits + 1
      ElseIf ch = ":" Then
         colon = i
      ElseIf Mid(TimeEntry, i, 1) = "a" Then
         ampm = "a"
         Exit For
      ElseIf Mid(TimeEntry, i, 1) = "p" Then
         ampm = "p"
         Exit For
      End If
   Next i
   
   If digits = 0 Then
      GetTime = ""
      Exit Function
   ElseIf colon = 0 Then
      Select Case digits
      Case 1
         ts = "0" & Mid(TimeEntry, 1, lastDigit) & ":00" & ampm
      Case 2
         ts = Mid(TimeEntry, 1, lastDigit) & ":00" & ampm
      Case 3
         ts = "0" & Mid(TimeEntry, 1, 1) & ":" & Mid(TimeEntry, 2, 2) & ampm
      Case 4
         ts = Mid(TimeEntry, 1, 2) & ":" & Mid(TimeEntry, 3, 2) & ampm
      Case Else
         GetTime = ""
         Exit Function
      End Select
   Else
      On Error Resume Next
      ts = Format(TimeEntry, "hh:mm a/p")
      ts = Replace(ts, " ", "")
      On Error GoTo 0
   End If
   
   GetTime = ts
   
End Function

Public Sub ComputeOverlappingCharges(empno As Long, CardDate As Date)
   
   'if no overlapped time charges, just set elapsed times
   
   Dim rdo As ADODB.Recordset
   Dim SQL As String
   Dim where As String
   where = "where TCEMP = " & empno & " and TMDAY = '" & Format(CardDate, "mm/dd/yyyy") & "'" & vbCrLf _
      & "and TCSURUN <> 'I'  AND TCSTOPTIME > TCSTARTTIME"
   
   SQL = _
      "select TCSTARTTIME as OPTIME, 'START' as OPSTARTSTOP, TCSURUN," & vbCrLf _
      & "TCPARTREF , TCRUNNO, TCOPNO" & vbCrLf _
      & "from TcitTable join TchdTable on TMCARD = TCCARD" & vbCrLf _
      & where & vbCrLf _
      & "Union ALL " & vbCrLf _
      & "select TCSTOPTIME as OPTIME, 'STOP' as OPSTARTSTOP, TCSURUN," & vbCrLf _
      & "TCPARTREF , TCRUNNO, TCOPNO" & vbCrLf _
      & "from TcitTable join TchdTable on TMCARD = TCCARD" & vbCrLf _
      & where & vbCrLf _
      & "ORDER BY OPTIME, OPSTARTSTOP DESC"
   
   Dim segments As Integer, jobCount As Integer
   segments = 0
   jobCount = 0
   Debug.Print SQL
   
   ' MM 9/8 clsADOCon.BeginTrans
   sSql = SQL
   If clsADOCon.GetDataSet(sSql, rdo) Then
      While Not rdo.EOF
         If rdo!OPSTARTSTOP = "START" Then
            segments = segments + 1
            If jobCount < segments Then
               jobCount = segments
            End If
         Else
            segments = segments - 1
         End If
         rdo.MoveNext
      Wend
   End If
   
   'calculate full segment elapsed times
   sSql = "Update TcitTable set TCHOURS = cast(cast(datediff(n, TCSTARTTIME, TCSTOPTIME) as decimal(10,6))" & vbCrLf _
      & "/ 60.0  as decimal(10,6))," & vbCrLf _
      & "TCMULTIJOB = 0" & vbCrLf _
      & "from TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & where
      Debug.Print sSql
   clsADOCon.ExecuteSql sSql
   
   'if no overlapping segments, we're done
   If jobCount <= 1 Then
      ' MM 9/8 clsADOCon.CommitTrans
      Exit Sub
   End If
   
   'there are overlapping charges.  subtract shared amounts for each segment
   segments = 0
   jobCount = 0
   Dim prevSegmentStart As Date, segmentStart As Date
   sSql = SQL
   If clsADOCon.GetDataSet(sSql, rdo, ES_STATIC) Then
      Do While Not rdo.EOF
         segmentStart = rdo!OpTime
         If segments > 1 Then
            Dim Hours As Currency
            Hours = (segments - 1) * DateDiff("n", prevSegmentStart, segmentStart) / (60 * segments)
            
            If Hours > 0 Then
               sSql = "update TcitTable set TCHOURS = cast(TCHOURS - " & Hours & " as decimal(10,6))," & vbCrLf _
                  & "TCMULTIJOB = case when TCMULTIJOB < " & segments & " then " & segments & " else TCMULTIJOB end" & vbCrLf _
                  & "where TCEMP = " & empno & vbCrLf _
                  & "and TCSTARTTIME <= '" & prevSegmentStart & "'" & vbCrLf _
                  & "and TCSTOPTIME >= '" & segmentStart & "'"

'Debug.Print prevSegmentStart & " - " & segmentStart & " count " & segments & " subtract " & Hours & vbCrLf
'Debug.Print sSql & vbCrLf
      
               clsADOCon.ExecuteSql sSql
            End If
         End If
         If rdo!OPSTARTSTOP = "START" Then
            segments = segments + 1
            If jobCount < segments Then
               jobCount = segments
            End If
         Else
            segments = segments - 1
         End If
         prevSegmentStart = segmentStart
         rdo.MoveNext
      Loop
   End If

   ' MM 9/8 clsADOCon.CommitTrans
End Sub

Sub UpdateTimeCardTotals(EmployeeNumber As Long, CardDate As Date, Optional bCopyToOrg As Boolean = False)
   
   'update daily totals for timecard
   'TchdTable.TMREGHRS,TMOVTHRS,TMDBLHRS,TMSTART,TMSTOP

   sSql = "UpdateTimeCardTotals " & EmployeeNumber & ", '" & Format(CardDate, "mm/dd/yyyy") & "'"
   clsADOCon.ExecuteSql sSql ',rdExecDirect
   
   If (bCopyToOrg = True) Then
      ' Take a backof of the origncal start and end time.
      sSql = "UPDATE TchdTable set TMSTARTORIG = TMSTART, TMSTOPORIG = TMSTOP " & _
          " WHERE TMDAY = '" & Format(CardDate, "mm/dd/yyyy") & "' AND TMEMP = '" & CStr(EmployeeNumber) & "'"
      clsADOCon.ExecuteSql sSql ',rdExecDirect
   End If
End Sub


Function ApplySCOnlyToLunch(iEmpNum As Long, strCardDate As String, _
                        strTCBeg As String, strTCEnd As String, strTCRegHrs As String, _
                        strSftBeg As String, strSftEnd As String, strSfRegHrs As String)
    
   Dim RdoEmp As ADODB.Recordset
   Dim strTCCard As String
   Dim strLBeg As String
   Dim strLEnd As String
   Dim strLunBegDate As String
   Dim strLunEndDate As String
   Dim strTCBegDate As String
   Dim strTCEndDate As String
   
   On Error Resume Next
   'On Error GoTo DiaErr1
    
    
'    sSql = "SELECT TMCARD, ISNULL(SFLUNSTHR, 0) SFLUNSTHR, " & _
'            "ISNULL(SFLUNENHR, 0) SFLUNENHR FROM viewShiftCdEmployeeDetail, TchdTable " & _
'        "WHERE TMEMP = PREMNUMBER AND " & _
'            "PREMNUMBER = '" & CStr(iEmpNum) & "' AND " & _
'            "TMDAY = '" & strCardDate & "'"
    
    sSql = "SELECT TMCARD, ISNULL(c.SFLUNSTHR, 0) SFLUNSTHR, " & _
            "ISNULL(c.SFLUNENHR, 0) SFLUNENHR " & _
         "FROM  viewShiftCdEmployeeDetail b, TchdTable, SfcdTable c " & _
      "WHERE TMEMP = PREMNUMBER AND TMDAY = '" & strCardDate & "' " & _
            "AND PREMNUMBER = '" & CStr(iEmpNum) & "' " & _
            "AND " & _
         "(SELECT DISTINCT " & _
            "(CASE DATEPART(weekday, '" & strCardDate & "') " & _
            "WHEN 1 THEN ISNULL(SFREFSUN, SFREF) " & _
            "WHEN  2 THEN ISNULL(SFREFMON, SFREF) " & _
            "WHEN  3 THEN ISNULL(SFREFTUE, SFREF) " & _
            "WHEN  4 THEN ISNULL(SFREFWED, SFREF) " & _
            "WHEN  5 THEN ISNULL(SFREFTHU, SFREF) " & _
            "WHEN  6 THEN ISNULL(SFREFFRI, SFREF) " & _
            "WHEN  7 THEN ISNULL(SFREFSAT, SFREF) " & _
            "Else SFREF " & _
            "END) " & _
            "FROM sfempTable a " & _
      "WHERE a.PREMNUMBER = b.PREMNUMBER) " & _
       " = c.SFREF ORDER BY PREMNUMBER "
           
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
    If bSqlRows Then
        strTCCard = Trim(RdoEmp!TMCARD)
        strLBeg = Trim(RdoEmp!SFLUNSTHR)
        strLEnd = Trim(RdoEmp!SFLUNENHR)
        ClearResultSet RdoEmp
    Else
        Exit Function
    End If
    
   Dim dtTCStart As Date
   Dim dtTCEnd As Date
   dtTCStart = CDate(strCardDate & " " & strTCBeg)
   ' If the shift end time goes to next day, move to next day
   If (CDate(strTCBeg) > CDate(strTCEnd)) Then
     dtTCEnd = CDate(DateAdd("d", 1, CDate(strCardDate)) & " " & strTCEnd)
   Else
     dtTCEnd = CDate(strCardDate & " " & strTCEnd)
   End If
   
   If ((strLBeg <> "") And (strLEnd <> "")) Then
   
      If (CDate(strSftBeg) > CDate(strLBeg)) Then
         strLunBegDate = Format(DateAdd("d", 1, CDate(strCardDate)), "mm/dd/yy ") & strLBeg
         strLunEndDate = Format(DateAdd("d", 1, CDate(strCardDate)), "mm/dd/yy ") & strLEnd
      Else
         strLunBegDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLBeg
         strLunEndDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLEnd
      End If
   Else
      strLunBegDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLBeg
      strLunEndDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLEnd
   End If
           
    ' Commmented: checking above to see if the Lunch begin next day
   'strLunBegDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLBeg
   'strLunEndDate = Format(CDate(strCardDate), "mm/dd/yy ") & strLEnd
   
   strTCBegDate = Format(CDate(dtTCStart), "mm/dd/yy ") & strTCBeg
   strTCEndDate = Format(CDate(dtTCEnd), "mm/dd/yy ") & strTCEnd
    
    
   ' Adjust the checked shift times and lunch time
   Dim lLunMin As Long
   Dim lSftMin As Long
   
   AdjWithLunchStartEndTime strTCBegDate, strTCEndDate, lLunMin, strLunBegDate, strLunEndDate
   
   If (IsValidTime(strTCBegDate) And _
               IsValidTime(strTCBegDate)) Then
       
       lSftMin = DateDiff("n", strTCBegDate, strTCEndDate)
       strSfRegHrs = Format((lSftMin - lLunMin) / 60, "##0.00")
   Else
       strSfRegHrs = Format(0, "##0.00")
   End If
        
             
    sSql = "UPDATE TchdTable SET TMREGHRS = '" & strSfRegHrs & "' " & _
        " WHERE TMCARD = '" & strTCCard & "' AND TMEMP = '" & CStr(iEmpNum) & "'" & _
            " AND TMDAY = '" & strCardDate & "'"
    clsADOCon.ExecuteSql sSql ', rdExecDirect
        
        
    ' Adjust the Luch time
    AdjustLunchTime iEmpNum, strTCCard, strCardDate, strTCBeg, _
                                    strTCEnd, strLBeg, strLEnd
    ' Compute overlap time
    ComputeOverlappingChargesWithLunch iEmpNum, CDate(strCardDate)
        
End Function


Function ApplyShiftCode(iEmpNum As Long, strCardDate As String, _
                        strTCBeg As String, strTCEnd As String, strTCRegHrs As String, _
                        strSftBeg As String, strSftEnd As String, strSfRegHrs As String)
    
    
    
   Dim RdoEmp As ADODB.Recordset
   Dim strTCCard As String
   Dim strLBeg As String
   Dim strLEnd As String
   Dim strAdjMin As String
   
   On Error Resume Next
   'On Error GoTo DiaErr1
    
        
'   sSql = "SELECT TMCARD, ISNULL(SFLUNSTHR, 0) SFLUNSTHR, " & _
'           "ISNULL(SFLUNENHR, 0) SFLUNENHR, ISNULL(SFADJHR, 0) SFADJHR FROM viewShiftCdEmployeeDetail, TchdTable " & _
'       "WHERE TMEMP = PREMNUMBER AND " & _
'           "PREMNUMBER = '" & CStr(iEmpNum) & "' AND " & _
'           "TMDAY = '" & strCardDate & "'"
   
    sSql = "SELECT TMCARD, ISNULL(c.SFLUNSTHR, 0) SFLUNSTHR, " & _
            "ISNULL(c.SFLUNENHR, 0) SFLUNENHR, ISNULL(c.SFADJHR, 0) SFADJHR " & _
         "FROM  viewShiftCdEmployeeDetail b, TchdTable, SfcdTable c " & _
      "WHERE TMEMP = PREMNUMBER AND TMDAY = '" & strCardDate & "' " & _
            "AND PREMNUMBER = '" & CStr(iEmpNum) & "' " & _
            "AND " & _
         "(SELECT DISTINCT " & _
            "(CASE DATEPART(weekday, '" & strCardDate & "') " & _
            "WHEN 1 THEN ISNULL(SFREFSUN, SFREF) " & _
            "WHEN  2 THEN ISNULL(SFREFMON, SFREF) " & _
            "WHEN  3 THEN ISNULL(SFREFTUE, SFREF) " & _
            "WHEN  4 THEN ISNULL(SFREFWED, SFREF) " & _
            "WHEN  5 THEN ISNULL(SFREFTHU, SFREF) " & _
            "WHEN  6 THEN ISNULL(SFREFFRI, SFREF) " & _
            "WHEN  7 THEN ISNULL(SFREFSAT, SFREF) " & _
            "Else SFREF " & _
            "END) " & _
            "FROM sfempTable a " & _
      "WHERE a.PREMNUMBER = b.PREMNUMBER) " & _
       " = c.SFREF ORDER BY PREMNUMBER "
   
   Debug.Print sSql

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp, ES_FORWARD)
   If bSqlRows Then
       strTCCard = Trim(RdoEmp!TMCARD)
       strLBeg = Trim(RdoEmp!SFLUNSTHR)
       strLEnd = Trim(RdoEmp!SFLUNENHR)
       strAdjMin = Trim(RdoEmp!SFADJHR)
       
       ClearResultSet RdoEmp
   Else
       Exit Function
   End If
   
   ' MM 9/8 clsADOCon.BeginTrans
    
   sSql = "UPDATE TchdTable SET TMSTART = '" & strSftBeg & "'," & _
           " TMSTOP = '" & strSftEnd & "', TMREGHRS = '" & strSfRegHrs & "'," & _
           " TMSFCODEAPPLD = 1 " & _
       " WHERE TMCARD = '" & strTCCard & "' AND TMEMP = '" & CStr(iEmpNum) & "'" & _
           " AND TMDAY = '" & strCardDate & "'"
   
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql ', rdExecDirect
        
   Dim dtStart As Date
   Dim dtEnd As Date
   Dim dtSfAdjStart As Date
   Dim dtSfAdjEnd As Date
    
   dtStart = CDate(strCardDate & " " & strSftBeg)
    
   dtSfAdjStart = DateAdd("n", CInt(strAdjMin), CDate(dtStart))
    
    
   Dim where As String
   Dim where1 As String
    
    where = "WHERE TMCARD = '" & strTCCard & "' AND TCEMP = " & CStr(iEmpNum) & vbCrLf _
        & " and TMDAY = '" & Format(strCardDate, "mm/dd/yyyy") & "'" & vbCrLf _
       & " AND TCSTOPTIME > TCSTARTTIME"
        
        
    If (CDate(strTCBeg) <> CDate(strSftBeg)) Then
    
        sSql = "UPDATE TcitTable SET TCSTART = '" & strSftBeg & "'," & _
                " TCSTARTTIME = '" & Format(dtStart, "mm/dd/yy hh:mm") & "' " & _
               " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
              & where & " AND TCSTARTTIME <= '" & Format(dtSfAdjStart, "mm/dd/yy hh:mm") & "'"
              '& where & " AND TCSTART < '" & Format(dtStart, "mm/dd/yy hh:mm") & "' AND TCSTART < '" & strSftBeg & "'"
        Debug.Print sSql
        clsADOCon.ExecuteSql sSql ', rdExecDirect
    
      ' Update the TCEND time if it is within tolerence time
      ' only if it indirect time
      where1 = "WHERE TMCARD = '" & strTCCard & "' AND TCEMP = " & CStr(iEmpNum) & vbCrLf _
          & " and TMDAY = '" & Format(strCardDate, "mm/dd/yyyy") & "'" & vbCrLf
        
        sSql = "UPDATE TcitTable SET TCSTOP = '" & strSftBeg & "'," & _
                " TCSTOPTIME = '" & Format(dtStart, "mm/dd/yy hh:mm") & "' " & _
               " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
              & where & " AND TCSTOPTIME <= '" & Format(dtSfAdjStart, "mm/dd/yy hh:mm") & "'" & vbCrLf _
                  & " AND TCSURUN = 'I'"
              '& where & " AND TCSTART < '" & Format(dtStart, "mm/dd/yy hh:mm") & "' AND TCSTART < '" & strSftBeg & "'"
        Debug.Print sSql
        clsADOCon.ExecuteSql sSql ', rdExecDirect
    End If
    
    ' If the shift end time goes to next day, move to next day
    If (CDate(strSftBeg) > CDate(strSftEnd)) Then
      dtEnd = CDate(DateAdd("d", 1, CDate(strCardDate)) & " " & strSftEnd)
    Else
      dtEnd = CDate(strCardDate & " " & strSftEnd)
    End If
    dtSfAdjEnd = DateAdd("n", -CInt(strAdjMin), CDate(dtEnd))
    
    ' MM TODO: Soem time the TCSTOP is same as Shift End time
    ' But the TcitTable time is not.
    'If (CDate(strTCEnd) <> CDate(strSftEnd)) Then
        sSql = "UPDATE TcitTable SET TCSTOP = '" & strSftEnd & "'," & _
                " TCSTOPTIME = '" & Format(dtEnd, "mm/dd/yy hh:mm") & "' " & _
               " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
              & where & " AND TCSTOPTIME >= '" & Format(dtSfAdjEnd, "mm/dd/yy hh:mm") & "'"
        
        'Debug.Print sSql
        clsADOCon.ExecuteSql sSql ', rdExecDirect
             
        ' If the job overlaps with the indirect time truncate
        'the start time shift end time
         where1 = "WHERE TMCARD = '" & strTCCard & "' AND TCEMP = " & CStr(iEmpNum) & vbCrLf _
             & " and TMDAY = '" & Format(strCardDate, "mm/dd/yyyy") & "'" & vbCrLf
        
        sSql = "UPDATE TcitTable SET TCSTART = '" & strSftEnd & "'," & _
                " TCSTARTTIME = '" & Format(dtEnd, "mm/dd/yy hh:mm") & "' " & _
               " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
              & where1 & " AND TCSTARTTIME >= '" & Format(dtSfAdjEnd, "mm/dd/yy hh:mm") & "'" & vbCrLf _
                  & " AND TCSURUN = 'I'"
        
        'Debug.Print sSql
        clsADOCon.ExecuteSql sSql ', rdExecDirect
    
    'End If
    ' MM 9/8 clsADOCon.CommitTrans
    
    ' Adjust the Luch time
    AdjustLunchTime iEmpNum, strTCCard, strCardDate, strSftBeg, _
                                    strSftEnd, strLBeg, strLEnd
    ' Compute overlap time
    ComputeOverlappingChargesWithLunch iEmpNum, CDate(strCardDate)
        
End Function

Public Sub AdjustLunchTime(iEmpNum As Long, strTCCard As String, strCardDate As String, _
                 strSftBeg As String, strSftEnd As String, strLStart As String, strLEnd As String)

    ' First adjust the Time Charge header data
    ' Get Lunch Minutes
    Dim rdo As ADODB.Recordset
    Dim lMinutes As Long
    Dim lHrs As Double
    
    ' Adjust the TC IT records
    
   sSql = "SELECT TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf _
      & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf _
      & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf _
      & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf _
      & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS " & vbCrLf _
      & " FROM TcitTable WHERE TCEMP = '" & CStr(iEmpNum) & "' AND " & _
            " TCCARD = '" & strTCCard & "'"
   Debug.Print sSql
    
    If clsADOCon.GetDataSet(sSql, rdo) Then
      While Not rdo.EOF
        Dim strTCBeg As String
        Dim strTCEnd As String
        Dim strLAdjBeg As String
        Dim strLAdjEnd As String
        Dim dtSfBeg As Date
        Dim dtSfEnd As Date
        Dim dtTCStart As Date
        Dim dtTCStop As Date
        Dim dtLunBeg As Date
        Dim dtLunEnd As Date
        Dim dtLAdjBeg As Date
        Dim dtLAdjEnd As Date
        Dim Comments As String
        Dim ret  As Integer
        
        Dim DebitAccount As String, DebitAccount1 As String, Shop As String, WorkCenter As String
        Dim Rate As Currency, OverheadRate As Currency, OverheadFixed As Currency

        Dim TimeCode As String, SetupRunIndirect As String, IndirectAccount As String, _
        PartNumber As String, RunNumber As Long, OpNumber As Integer, _
        JournalRef As Integer, Source As String, _
        AcceptCount As Integer, RejectCount As Integer, ScrapCount As Integer
        Dim sortKey As Long
        Dim JournalID As String


        strTCBeg = Trim(rdo!TCSTART)
        strTCEnd = Trim(rdo!TCSTOP)
        
        TimeCode = Trim(rdo!TCCODE)
        Rate = Trim(rdo!TCRATE)
        OverheadRate = Trim(rdo!TCOHRATE)
        DebitAccount = Trim(rdo!TCACCT)
        DebitAccount1 = Trim(rdo!TCACCOUNT)
        Shop = Trim(rdo!TCSHOP)
        WorkCenter = Trim(rdo!TCWC)
        SetupRunIndirect = Trim(rdo!TCSURUN)
        PartNumber = Trim(rdo!TCPARTREF)
        RunNumber = Trim(rdo!TCRUNNO)
        OpNumber = Trim(rdo!TCOPNO)
        sortKey = Trim(rdo!TCSORT)
        OverheadFixed = Trim(rdo!TCOHFIXED)
        JournalID = Trim(rdo!TCGLJOURNAL)
        JournalRef = Trim(rdo!TCGLREF)
        Source = Trim(rdo!TCSOURCE)
        AcceptCount = Trim(rdo!TCACCEPT)
        RejectCount = Trim(rdo!TCREJECT)
        ScrapCount = Trim(rdo!TCSCRAP)
        dtTCStart = Trim(rdo!TCSTARTTIME)
        dtTCStop = Trim(rdo!TCSTOPTIME)
        Comments = Replace(Trim(rdo!TCCOMMENTS), "'", "''")
        
         If (CDate(strSftBeg) > CDate(strSftEnd)) Then
            dtSfBeg = Format(CDate(strCardDate), "mm/dd/yy ") & strSftBeg
            dtSfEnd = Format(DateAdd("d", 1, CDate(strCardDate)), "mm/dd/yy ") & strSftEnd
         Else
            dtSfBeg = Format(CDate(strCardDate), "mm/dd/yy ") & strSftBeg
            dtSfEnd = Format(CDate(strCardDate), "mm/dd/yy ") & strSftEnd
         End If
         
         

         If ((strLStart <> "") And (strLEnd <> "")) Then
         
            If (CDate(strSftBeg) > CDate(strLStart)) Then
               dtLunBeg = Format(DateAdd("d", 1, CDate(strCardDate)), "mm/dd/yy ") & strLStart
               dtLunEnd = Format(DateAdd("d", 1, CDate(strCardDate)), "mm/dd/yy ") & strLEnd
            Else
               dtLunBeg = Format(CDate(strCardDate), "mm/dd/yy ") & strLStart
               dtLunEnd = Format(CDate(strCardDate), "mm/dd/yy ") & strLEnd
            End If
            'ret = GetLunchPeriod(strTCBeg, strTCEnd, strLStart, strLEnd, strLAdjBeg, strLAdjEnd)
            ret = GetLunchPeriod(CStr(dtTCStart), CStr(dtTCStop), CStr(dtLunBeg), CStr(dtLunEnd), strLAdjBeg, strLAdjEnd)
         Else
            ret = 0
         End If
        
        If (ret > 0) Then
            
            dtLAdjBeg = CDate(strLAdjBeg)
            dtLAdjEnd = CDate(strLAdjEnd)
            If ((CDate(dtLAdjBeg) > CDate(dtTCStart)) And _
                (CDate(dtLAdjEnd) < CDate(dtTCStop))) Then
                
                Dim dtStart As Date
                Dim dtEnd As Date
                Dim lHours As Currency
                Dim hoursAsDateTime As String
                Dim StartDateTimeToMinute As Variant, EndDateTimeToMinute As Variant
                
                'dtStart = CDate(strCardDate & " " & strTCBeg)
                'dtEnd = CDate(strCardDate & " " & strLAdjBeg)
                dtStart = CDate(dtTCStart)
                dtEnd = CDate(dtLAdjBeg)
                StartDateTimeToMinute = Format(dtStart, "mm/dd/yy hh:mm")
                EndDateTimeToMinute = Format(dtEnd, "mm/dd/yy hh:mm")
                
                ComputeElapsedTime StartDateTimeToMinute, EndDateTimeToMinute, lHours
                hoursAsDateTime = ConvertHours(lHours)
                    
                If lHours > 0 Then
                     sSql = "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf _
                        & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf _
                        & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf _
                        & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf _
                        & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf _
                        & "VALUES('" & strTCCard & "'," & CStr(iEmpNum) & ",'" & GetTime(dtTCStart) & "'," & vbCrLf _
                        & "'" & GetTime(dtLAdjBeg) & "'," & "'" & StartDateTimeToMinute & "', '" & EndDateTimeToMinute & "'," & vbCrLf _
                        & CStr(lHours) & ",'" & hoursAsDateTime & "'," & vbCrLf _
                        & "'" & TimeCode & "'," & Rate & "," & OverheadRate & "," & vbCrLf _
                        & "1,'" & DebitAccount & "','" & DebitAccount1 & "','" & Shop & "','" & WorkCenter & "'," _
                        & "0,'" & SetupRunIndirect & "',0,'" & PartNumber & "'," _
                        & RunNumber & "," & OpNumber & "," & sortKey & "," & OverheadFixed & "," & vbCrLf _
                        & "'" & JournalID & "'," & JournalRef & ",'" & Source & "'," & vbCrLf _
                        & "0, " & AcceptCount & ", " & RejectCount & ", " & ScrapCount & ",'" & Comments & "')"
                  
                     clsADOCon.ExecuteSql sSql
                End If
                
                'dtStart = CDate(strCardDate & " " & strLAdjEnd)
                'dtEnd = CDate(strCardDate & " " & strTCEnd)
                dtStart = CDate(dtLAdjEnd)
                dtEnd = CDate(dtTCStop)
                StartDateTimeToMinute = Format(dtStart, "mm/dd/yy hh:mm")
                EndDateTimeToMinute = Format(dtEnd, "mm/dd/yy hh:mm")
                
                ComputeElapsedTime StartDateTimeToMinute, EndDateTimeToMinute, lHours
                hoursAsDateTime = ConvertHours(lHours)
                
               If lHours > 0 Then
                  sSql = "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf _
                     & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf _
                     & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf _
                     & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf _
                     & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP,TCCOMMENTS)" & vbCrLf _
                     & "VALUES('" & strTCCard & "'," & CStr(iEmpNum) & ",'" & GetTime(dtLAdjEnd) & "'," & vbCrLf _
                     & "'" & GetTime(dtTCStop) & "'," & "'" & StartDateTimeToMinute & "', '" & EndDateTimeToMinute & "'," & vbCrLf _
                     & CStr(lHours) & ",'" & hoursAsDateTime & "'," & vbCrLf _
                     & "'" & TimeCode & "'," & Rate & "," & OverheadRate & "," & vbCrLf _
                     & "1,'" & DebitAccount & "','" & DebitAccount1 & "','" & Shop & "','" & WorkCenter & "'," _
                     & "0,'" & SetupRunIndirect & "',0,'" & PartNumber & "'," _
                     & RunNumber & "," & OpNumber & "," & sortKey & "," & OverheadFixed & "," & vbCrLf _
                     & "'" & JournalID & "'," & JournalRef & ",'" & Source & "'," & vbCrLf _
                     & "0, " & AcceptCount & ", " & RejectCount & ", " & ScrapCount & ",'" & Comments & "')"
                  
                  clsADOCon.ExecuteSql sSql
               End If
                sSql = "DELETE FROM TcitTable WHERE TCCARD = '" & strTCCard & "' AND TCEMP = '" & CStr(iEmpNum) & "'" & vbCrLf _
                   & " AND TCSTART = '" & strTCBeg & "' AND TCSTOP = '" & strTCEnd & "' AND " & _
                        "TCSORT = '" & CStr(sortKey) & "'"
                
                clsADOCon.ExecuteSql sSql
            
            End If
            
            If (CDate(dtLAdjBeg) = CDate(dtTCStart)) Then
                If (CDate(dtTCStop) >= CDate(dtLAdjEnd)) Then
                    
                    'dtStart = CDate(strCardDate & " " & strLAdjEnd)
                    'dtEnd = CDate(strCardDate & " " & strTCEnd)
                    dtStart = CDate(dtLAdjEnd)
                    dtEnd = CDate(dtTCStop)
                    StartDateTimeToMinute = Format(dtStart, "mm/dd/yy hh:mm")
                    EndDateTimeToMinute = Format(dtEnd, "mm/dd/yy hh:mm")
                    
                    ComputeElapsedTime StartDateTimeToMinute, EndDateTimeToMinute, lHours
                    hoursAsDateTime = ConvertHours(lHours)
                    
                    sSql = "Update TcitTable set TCSTART = '" & GetTime(dtLAdjEnd) & "', " & _
                            " TCHOURS = '" & CStr(lHours) & "'," & _
                            " TCSTARTTIME = '" & StartDateTimeToMinute & "'," & _
                            " TCMULTIJOB = 0 " & _
                       " FROM TcitTable WHERE TCCARD = '" & strTCCard & "' AND " & _
                            " TCEMP = '" & CStr(iEmpNum) & "'" & _
                            " AND TCSTART = '" & strTCBeg & "' AND TCSTOP = '" & strTCEnd & "'"
                        'strBeg = strEnd1
                    clsADOCon.ExecuteSql sSql
                End If
            End If
            
            If (CDate(dtLAdjEnd) = CDate(dtTCStop)) Then
                If (CDate(dtLAdjBeg) > CDate(dtTCStart)) Then
                    
                    'dtStart = CDate(strCardDate & " " & strTCBeg)
                    'dtEnd = CDate(strCardDate & " " & strLAdjBeg)
                    dtStart = CDate(dtTCStart)
                    dtEnd = CDate(dtLAdjBeg)
                    StartDateTimeToMinute = Format(dtStart, "mm/dd/yy hh:mm")
                    EndDateTimeToMinute = Format(dtEnd, "mm/dd/yy hh:mm")
                    
                    ComputeElapsedTime StartDateTimeToMinute, EndDateTimeToMinute, lHours
                    hoursAsDateTime = ConvertHours(lHours)
                    
                    sSql = "Update TcitTable set TCSTOP = '" & GetTime(dtLAdjBeg) & "', " & _
                            " TCHOURS = '" & CStr(lHours) & "'," & _
                            " TCSTOPTIME = '" & EndDateTimeToMinute & "'," & _
                            " TCMULTIJOB = 0 " & _
                       " FROM TcitTable WHERE TCCARD = '" & strTCCard & "' AND " & _
                            " TCEMP = '" & CStr(iEmpNum) & "'" & _
                            " AND TCSTART = '" & strTCBeg & "' AND TCSTOP = '" & strTCEnd & "'"
                    clsADOCon.ExecuteSql sSql
                    
                    'strEnd = strBeg1
                End If
            End If
            
        Else
        
            'dtStart = CDate(strCardDate & " " & strTCBeg)
            'dtEnd = CDate(strCardDate & " " & strTCEnd)
            Dim strUpdateStopTime As String
            dtStart = CDate(dtTCStart)
            dtEnd = CDate(dtTCStop)
            ' This is because the Stop could be less than adjusted shift start time.
            If (CDate(dtTCStart) > CDate(dtEnd)) Then
               strUpdateStopTime = ", TCSTOP = '" & GetTime(dtTCStart) & "'"
               dtEnd = CDate(dtTCStart)
            Else
               strUpdateStopTime = ""
               dtEnd = CDate(dtTCStop)
            End If

            
            StartDateTimeToMinute = Format(dtStart, "mm/dd/yy hh:mm")
            EndDateTimeToMinute = Format(dtEnd, "mm/dd/yy hh:mm")
            
            ComputeElapsedTime StartDateTimeToMinute, EndDateTimeToMinute, lHours
            hoursAsDateTime = ConvertHours(lHours)
        
            sSql = "Update TcitTable SET TCHOURS = '" & CStr(lHours) & "'," & _
                    " TCMULTIJOB = 0 " & strUpdateStopTime & _
               " FROM TcitTable WHERE TCCARD = '" & strTCCard & "' AND " & _
                    " TCEMP = '" & CStr(iEmpNum) & "'" & _
                    " AND TCSTART = '" & GetTime(dtTCStart) & "' AND TCSTOP = '" & GetTime(dtTCStop) & "'"
            clsADOCon.ExecuteSql sSql
        End If ' Lunch period
            
      rdo.MoveNext
      Wend
    End If ' record set
    
End Sub
Public Sub ComputeOverlappingChargesWithLunch(empno As Long, CardDate As Date)
                                
   
   'if no overlapped time charges, just set elapsed times
   
   Dim rdo As ADODB.Recordset
   Dim SQL As String
   Dim where As String
   where = "where TCEMP = " & empno & " and TMDAY = '" & Format(CardDate, "mm/dd/yyyy") & "'" & vbCrLf _
      & "and TCSURUN <> 'I'  AND TCSTOPTIME > TCSTARTTIME"
   
   SQL = _
      "select TCSTARTTIME as OPTIME, 'START' as OPSTARTSTOP, TCSURUN," & vbCrLf _
      & "TCPARTREF , TCRUNNO, TCOPNO" & vbCrLf _
      & "from TcitTable join TchdTable on TMCARD = TCCARD" & vbCrLf _
      & where & vbCrLf _
      & "Union" & vbCrLf _
      & "select TCSTOPTIME as OPTIME, 'STOP' as OPSTARTSTOP, TCSURUN," & vbCrLf _
      & "TCPARTREF , TCRUNNO, TCOPNO" & vbCrLf _
      & "from TcitTable join TchdTable on TMCARD = TCCARD" & vbCrLf _
      & where & vbCrLf _
      & "ORDER BY OPTIME, OPSTARTSTOP DESC"
    
   Debug.Print SQL
   
   Dim segments As Integer, jobCount As Integer
   segments = 0
   jobCount = 0
   
   ' MM 9/8 clsADOCon.BeginTrans
   sSql = SQL
   If clsADOCon.GetDataSet(sSql, rdo) Then
      While Not rdo.EOF
         If rdo!OPSTARTSTOP = "START" Then
            segments = segments + 1
            If jobCount < segments Then
               jobCount = segments
            End If
         Else
            segments = segments - 1
         End If
         rdo.MoveNext
      Wend
   End If
   
   
   'if no overlapping segments, we're done
   If jobCount <= 1 Then
      ' MM 9/8 clsADOCon.CommitTrans
      Exit Sub
   End If
   
   'there are overlapping charges.  subtract shared amounts for each segment
   segments = 0
   jobCount = 0
   Dim prevSegmentStart As Date, segmentStart As Date
   sSql = SQL
   If clsADOCon.GetDataSet(sSql, rdo, ES_STATIC) Then
      Do While Not rdo.EOF
         segmentStart = rdo!OpTime
         If segments > 1 Then
            Dim Hours As Currency
            ' Get the Lunch minutes
            Dim strSegBeg, strSegEnd As String
            strSegBeg = Format(prevSegmentStart, "hh:mma/p")
            strSegEnd = Format(segmentStart, "hh:mma/p")
            
            Hours = (segments - 1) * (DateDiff("n", prevSegmentStart, segmentStart)) / (60 * segments)
            If Hours > 0 Then
               sSql = "update TcitTable set TCHOURS = cast(TCHOURS - " & Hours & " as decimal(10,3))," & vbCrLf _
                  & "TCMULTIJOB = case when TCMULTIJOB < " & segments & " then " & segments & " else TCMULTIJOB end" & vbCrLf _
                  & "where TCEMP = " & empno & vbCrLf _
                  & "and TCSTARTTIME <= '" & prevSegmentStart & "'" & vbCrLf _
                  & "and TCSTOPTIME >= '" & segmentStart & "'" & vbCrLf _
                  & "and TCSURUN <> 'I'"
'Debug.Print prevSegmentStart & " - " & segmentStart & " count " & segments & " subtract " & Hours
               Debug.Print sSql
               clsADOCon.ExecuteSql sSql
            End If
         End If
         If rdo!OPSTARTSTOP = "START" Then
            segments = segments + 1
            If jobCount < segments Then
               jobCount = segments
            End If
         Else
            segments = segments - 1
         End If
         prevSegmentStart = segmentStart
         rdo.MoveNext
      Loop
   End If

   ' MM 9/8 clsADOCon.CommitTrans
End Sub

Function GetLunchPeriod(strBeg As String, strEnd As String, strLStart As String, _
                        strLEnd As String, ByRef strLAdjBeg As String, ByRef strLAdjEnd As String)
    
    Dim strBeg1, strEnd1 As String
    Dim lLHours As Long
    
    strLAdjEnd = ""
    strLAdjBeg = ""
    If (CDate(strLStart) >= CDate(strEnd)) Then
        GetLunchPeriod = 0
        Exit Function
    End If
    If (CDate(strLEnd) <= CDate(strBeg)) Then
        GetLunchPeriod = 0
        Exit Function
    End If
    
    If (CDate(strLStart) >= CDate(strBeg)) Then
        strLAdjBeg = strLStart
    Else
        strLAdjBeg = strBeg
    End If
    
    If (CDate(strLEnd) >= CDate(strEnd)) Then
        strLAdjEnd = strEnd
    Else
        strLAdjEnd = strLEnd
    End If
    
    GetLunchPeriod = 1
    
End Function

Function AdjustShiftStartEndTime(ByRef strBeg As String, ByRef strEnd As String, _
                                    ByVal strSfBeg As String, ByVal strSfEnd As String, _
                                    ByVal strAdjMin As String)
    
    If (CDate(strSfBeg) >= CDate(strBeg)) Then
        strBeg = strSfBeg
    ElseIf (CDate(strSfBeg) >= DateAdd("n", -CInt(strAdjMin), (CDate(strBeg)))) Then
        strBeg = strSfBeg
    End If
    
    If (CDate(strSfEnd) <= CDate(strEnd)) Then
        strEnd = strSfEnd
    ElseIf (CDate(strSfEnd) <= DateAdd("n", CInt(strAdjMin), (CDate(strEnd)))) Then
        strEnd = strSfEnd
    End If
    

End Function

Function AdjWithLunchStartEndTime(ByRef strBeg As String, ByRef strEnd As String, ByRef lMinutes As Long, _
                                    ByVal strLBeg As String, ByVal strLEnd As String) As Boolean

    Dim strBeg1 As String
    Dim strEnd1 As String
    
    lMinutes = 0
    If (CDate(strLBeg) >= CDate(strEnd)) Then
        AdjWithLunchStartEndTime = True
        Exit Function
    End If
    
    If (CDate(strLEnd) <= CDate(strBeg)) Then
        AdjWithLunchStartEndTime = True
        Exit Function
    End If
    
    If (CDate(strLBeg) >= CDate(strBeg)) Then
        strBeg1 = strLBeg
    Else
        strBeg1 = strBeg
    End If
    
    If (CDate(strLEnd) >= CDate(strEnd)) Then
        strEnd1 = strEnd
    Else
        strEnd1 = strLEnd
    End If
    
    ' if the checked time falls with in the lunch time
    ' then the lunch end time is same was shift start time
    If (CDate(strBeg1) = CDate(strBeg)) Then
        If (CDate(strEnd) >= CDate(strEnd1)) Then
            strBeg = strEnd1
        End If
    End If
    
    If (CDate(strEnd1) = CDate(strEnd)) Then
        If (CDate(strBeg1) > CDate(strBeg)) Then
            strEnd = strBeg1
        End If
    End If
    
    If ((CDate(strBeg1) > CDate(strBeg)) And _
        (CDate(strEnd1) < CDate(strEnd))) Then
        lMinutes = Abs(DateDiff("n", strBeg1, strEnd1))
    End If
    
    AdjWithLunchStartEndTime = True
    
End Function


Public Function GetShiftStartDate(empno As Long, dtstartJob As Date) As String

   Dim rdo As ADODB.Recordset
   Dim strSftBeg As String
   Dim strSftEnd As String
   Dim dtSftBeg As Date
   Dim dtSftEnd As Date
   Dim strDefScode As String
   
   sSql = "SELECT ISNULL(SFSTHR, '') SFSTHR, ISNULL(SFENHR, '') SFENHR " & vbCrLf _
            & ",ISNULL(SFOVERLAP, 0) SFOVERLAP, a.SFREF SFREF FROM dbo.sfempTable a,sfcdTable b WHERE " & vbCrLf _
            & " a.SFREF = b.SFREF And premnumber = '" & CStr(empno) & "'"
         
    Debug.Print sSql
    
    If clsADOCon.GetDataSet(sSql, rdo) Then
      If ((Not rdo.EOF) And (rdo!SFOVERLAP = 1)) Then
         strSftBeg = Trim(rdo!SFSTHR)
         strSftEnd = Trim(rdo!SFENHR)
         strDefScode = Trim(rdo!SFREF)
         
         ' Correct the shift code if we have alternate
         CheckForAltShift dtstartJob, CStr(empno), strDefScode, strSftBeg, strSftEnd
         
         ' adjust the shift start date and end time
         GetShiftDateTime strSftBeg, strSftEnd, dtstartJob, dtSftBeg, dtSftEnd
         
         Dim dateString As String
         dateString = Format(dtSftBeg, "mm/dd/yyyy")
         GetShiftStartDate = dateString
      Else
         GetShiftStartDate = ""
      End If
      rdo.Close
   Else
      GetShiftStartDate = ""
   End If
End Function

Public Function GetShiftDateTime(ByVal strSftBeg As String, ByVal strSftEnd As String, _
            ByVal dtstartJob As Date, ByRef dtSftBeg As Date, _
               ByRef dtSftEnd As Date)

   If (CDate(strSftBeg) > CDate(strSftEnd)) Then
      dtSftBeg = Format(CDate(dtstartJob), "mm/dd/yy ") & strSftBeg
      dtSftEnd = Format(DateAdd("d", 1, CDate(dtstartJob)), "mm/dd/yy ") & strSftEnd
   Else
      dtSftBeg = Format(CDate(dtstartJob), "mm/dd/yy ") & strSftBeg
      dtSftEnd = Format(CDate(dtstartJob), "mm/dd/yy ") & strSftEnd
   End If
      
   ' check if the job is with in the shift hours.
   ' Reduce the housrs by 2 so that the TC will be charged for today
   If (DateAdd("h", -4, CDate(dtSftBeg)) > CDate(dtstartJob)) Then
      dtSftBeg = DateAdd("d", -1, CDate(dtSftBeg))
      dtSftEnd = DateAdd("d", -1, CDate(dtSftEnd))
   End If

End Function

Public Function CheckForAltShift(dtstartJob As Date, strEmpNo As String, _
                  strDefScode As String, ByRef strSftBeg As String, _
                  ByRef strSftEnd As String)
                  
   Dim rdo As ADODB.Recordset
   Dim arWkDayShift(7) As String
   
   sSql = "SELECT SFREF, SFREFSUN, SFREFMON, SFREFTUE,SFREFWED," & vbCrLf _
             & "SFREFTHU, SFREFFRI,SFREFSAT " & vbCrLf _
         & " FROM SfempTable " & vbCrLf _
         & " WHERE PREMNUMBER = '" & strEmpNo & "'" & vbCrLf _
            & " AND SFREF = '" & strDefScode & "'"
    If clsADOCon.GetDataSet(sSql, rdo) Then
      If (Not rdo.EOF) Then
         
         Dim strTmp As String
         strTmp = (IIf(IsNull(rdo!SFREFSUN), rdo!SFREF, rdo!SFREFSUN))
         arWkDayShift(0) = strTmp
         arWkDayShift(1) = IIf(IsNull(rdo!SFREFMON), rdo!SFREF, rdo!SFREFMON)
         arWkDayShift(2) = IIf(IsNull(rdo!SFREFTUE), rdo!SFREF, rdo!SFREFTUE)
         arWkDayShift(3) = IIf(IsNull(rdo!SFREFWED), rdo!SFREF, rdo!SFREFWED)
         arWkDayShift(4) = IIf(IsNull(rdo!SFREFTHU), rdo!SFREF, rdo!SFREFTHU)
         arWkDayShift(5) = IIf(IsNull(rdo!SFREFFRI), rdo!SFREF, rdo!SFREFFRI)
         arWkDayShift(6) = IIf(IsNull(rdo!SFREFSAT), rdo!SFREF, rdo!SFREFSAT)
      End If
      rdo.Close
   End If
   Set rdo = Nothing
   
   Dim strWkdaySf As String
   Dim strNewSftBeg As String
   Dim strNewSftEnd As String
   Dim dtNewSftBeg As Date
   
   strWkdaySf = arWkDayShift((Weekday(dtstartJob) - 1))
   
   If (strWkdaySf <> strDefScode) Then
      Dim RdoNew As ADODB.Recordset
      sSql = "SELECT ISNULL(SFSTHR, '') SFSTHR, ISNULL(SFENHR, '') SFENHR " & vbCrLf _
               & " FROM sfcdTable WHERE sfref = '" & strWkdaySf & "'"

       If clsADOCon.GetDataSet(sSql, RdoNew) Then
         If (Not RdoNew.EOF) Then
            strNewSftBeg = Trim(RdoNew!SFSTHR)
            strNewSftEnd = Trim(RdoNew!SFENHR)
         End If
         RdoNew.Close
      End If
      Set RdoNew = Nothing
      
      ' if the new Shift
      If ((strNewSftBeg <> "") And (strNewSftEnd <> "")) Then
         dtNewSftBeg = Format(CDate(dtstartJob), "mm/dd/yy ") & strNewSftBeg
         
         If (CDate(dtstartJob) > DateAdd("h", -4, CDate(dtNewSftBeg))) Then
            strSftBeg = strNewSftBeg
            strSftEnd = strNewSftEnd
         End If
      End If
   End If
   
End Function
                           
Public Function GetTCardDates(empno As Long, ByVal strSelDate As String, _
            ByVal strTCBeg As String, ByVal strTCEnd As String, _
            ByRef dtTCStart As Date, ByRef dtTCEnd As Date)

   Dim rdo As ADODB.Recordset
   Dim strSftBeg As String
   Dim strSftEnd As String
   Dim dtSftBeg As Date
   Dim dtSftEnd As Date
   Dim strDefScode As String
   
   If (CDate(strTCBeg) > CDate(strTCEnd)) Then
      dtTCStart = Format(CDate(strSelDate), "mm/dd/yy ") & strTCBeg
      dtTCEnd = Format(DateAdd("d", 1, CDate(strSelDate)), "mm/dd/yy ") & strTCEnd
      
   Else
      dtTCStart = Format(CDate(strSelDate), "mm/dd/yy ") & strTCBeg
      dtTCEnd = Format(CDate(strSelDate), "mm/dd/yy ") & strTCEnd
   End If
   
   sSql = "SELECT a.SFREF, ISNULL(SFSTHR, '') SFSTHR, ISNULL(SFENHR, '') SFENHR " & vbCrLf _
            & ",ISNULL(SFOVERLAP, 0) SFOVERLAP FROM dbo.sfempTable a,sfcdTable b WHERE " & vbCrLf _
            & " a.SFREF = b.SFREF And premnumber = '" & CStr(empno) & "'"
         
    Debug.Print sSql
    
    
    If clsADOCon.GetDataSet(sSql, rdo) Then
      If ((Not rdo.EOF) And (rdo!SFOVERLAP = 1)) Then
         strSftBeg = Trim(rdo!SFSTHR)
         strSftEnd = Trim(rdo!SFENHR)
         strDefScode = Trim(rdo!SFREF)
         
         ' Correct the shift code if we have alternate
         CheckForAltShift dtTCStart, CStr(empno), strDefScode, strSftBeg, strSftEnd
         
         ' check if the job is with in the shift hours.
         ' Reduce the housrs by 2 so that the TC will be charged for today
         If (DateAdd("h", -4, CDate(strSftBeg)) > CDate(strTCBeg)) Then
            dtTCStart = DateAdd("d", 1, CDate(dtTCStart))
            dtTCEnd = DateAdd("d", 1, CDate(dtTCEnd))
         End If
         
         rdo.Close
      End If
   End If
   Set rdo = Nothing
   
End Function

