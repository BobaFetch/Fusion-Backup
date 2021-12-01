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

Public Function CreateTimeCharge(TimeCardID As String, EmpNo As Long, _
   StartDateTime As Variant, EndDateTime As Variant, TimeCode As String, _
   SetupRunIndirect As String, IndirectAccount As String, _
   PartNumber As String, RunNumber As Long, OpNumber As Integer, _
   JournalRef As Integer, Source As String, _
   AcceptCount, RejectCount, ScrapCount) As Boolean
   
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
      GetEmployeeAccountAndRate EmpNo, TimeCode, DebitAccount, Rate
      DebitAccount = IndirectAccount      'use account provided
   Else
      GetOpAccountAndRates PartNumber, RunNumber, OpNumber, EmpNo, TimeCode, DebitAccount, Rate, _
         OverheadFixed, OverheadRate, Shop, WorkCenter
   End If

   If DebitAccount = "" Then
      DebitAccount = GetDefTimeAccounts("Time")
   End If
   
   
   Dim startString As String, endString As String
   startString = GetTime(StartDateTime)
   endString = GetTime(EndDateTime)
   
   sSql = "INSERT INTO TcitTable (TCCARD,TCEMP,TCSTART,TCSTOP,TCSTARTTIME,TCSTOPTIME," & vbCrLf _
      & "TCHOURS,TCTIME,TCCODE,TCRATE,TCOHRATE,TCRATENO,TCACCT,TCACCOUNT," & vbCrLf _
      & "TCSHOP,TCWC,TCPAYTYPE,TCSURUN,TCYIELD,TCPARTREF,TCRUNNO," & vbCrLf _
      & "TCOPNO,TCSORT,TCOHFIXED,TCGLJOURNAL,TCGLREF,TCSOURCE," & vbCrLf _
      & "TCMULTIJOB,TCACCEPT,TCREJECT,TCSCRAP)" & vbCrLf _
      & "VALUES('" & TimeCardID & "'," & EmpNo & ",'" & startString & "'," & vbCrLf _
      & "'" & endString & "'," & "'" & StartDateTimeToMinute & "', '" & EndDateTimeToMinute & "'," & vbCrLf _
      & Hours & ",'" & hoursAsDateTime & "'," & vbCrLf _
      & "'" & TimeCode & "'," & Rate & "," & OverheadRate & "," & vbCrLf _
      & "1,'" & DebitAccount & "','" & DebitAccount & "','" & Shop & "','" & WorkCenter & "'," _
      & "0,'" & SetupRunIndirect & "',0,'" & PartNumber & "'," _
      & RunNumber & "," & OpNumber & "," & sortKey & "," & OverheadFixed & "," & vbCrLf _
      & "'" & JournalID & "'," & JournalRef & ",'" & Source & "'," & vbCrLf _
      & "0, " & AcceptCount & ", " & RejectCount & ", " & ScrapCount & ")"
      
'If SetupRunIndirect = "I" Then
'   Debug.Print startString & " - " & endString
'End If
   RdoCon.Execute sSql, rdExecDirect
   CreateTimeCharge = True
   Exit Function
   
whoops:
   sProcName = "CreateTimeCharge"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Dim frm As New ClassErrorForm
   DoModuleErrors frm

End Function

Public Sub GetOpAccountAndRates(PartNo As String, ByVal Runno As Long, _
   opNo As Integer, EmpNo As Long, TimeCode As String, Account As String, Rate As Currency, _
   FixedOhRate As Currency, PctOhRate As Currency, Shop As String, WorkCenter As String)
   
   'returned values:
   'Account (debit)
   'Rate
   'FixedOhRate
   'PctOhRate
   'Shop
   'WorkCenter
   
   Dim rdo As rdoResultset
'
'   'get employee information
'   Dim empRate As Currency, empAccount As String
'   sSql = "select PREMPAYRATE, rtrim(PREMACCTS) as PREMACCTS from EmplTable where PREMNUMBER = " & EmpNo
'   gblnSqlRows = GetDataSet(rdo)
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
   GetEmployeeAccountAndRate EmpNo, TimeCode, empAccount, empRate
   
   'get default information
   Dim companyAccount As String
   sSql = "select rtrim(WIPDRLABACCT) as WIPDRLABACCT from ComnTable"
   gblnSqlRows = GetDataSet(rdo)
   If gblnSqlRows Then
      companyAccount = rdo!WIPDRLABACCT
   Else
      companyAccount = ""
   End If
   rdo.Close
   
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
   gblnSqlRows = GetDataSet(rdo)
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
   End If
   
'   'multiple by rate for overtime, doubletime
'   Dim multiplier As Currency
'   multiplier = Me.GetTimeCodeMultiplier(TimeCode)
'   If multiplier <> 0 Then
'      Rate = multiplier * Rate
'   End If
'
End Sub

Public Sub GetEmployeeAccountAndRate(EmpNo As Long, TimeCode As String, empAccount As String, empRate As Currency)
   
   'get employee information
   Dim rdo As rdoResultset
   sSql = "select PREMPAYRATE, rtrim(PREMACCTS) as PREMACCTS from EmplTable where PREMNUMBER = " & EmpNo
   If GetDataSet(rdo) Then
      empRate = rdo!PREMPAYRATE * Me.GetTimeCodeMultiplier(TimeCode)
      empAccount = rdo!PREMACCTS
   Else
      empRate = 0
      empAccount = ""
   End If
End Sub

Public Function GetTimeCodeMultiplier(TimeCode As String) As Currency
   'return = 1 if no such timecode found or timecode has no multiplier
   
   Dim rdo As rdoResultset
   sSql = "SELECT TYPEADDER FROM TmcdTable WHERE TYPECODE = '" & TimeCode & "'"
   If GetDataSet(rdo) Then
      GetTimeCodeMultiplier = rdo!TYPEADDER
   End If
   
   If GetTimeCodeMultiplier = 0 Then
      GetTimeCodeMultiplier = 1
   End If
End Function

Public Function GetDefTimeAccounts(AccountType As String) As String
   'Time for Time Cards and Labor for distribution
   'AccountType = "Time" or "Labor"
   
   Dim RdoAcc As rdoResultset
   On Error GoTo modErr1
   sSql = "SELECT CODEFTIMEACCT,CODEFLABORACCT FROM ComnTable " _
          & "WHERE COREF=1"
   bSqlRows = GetDataSet(RdoAcc, ES_FORWARD)
   If bSqlRows Then
      With RdoAcc
         If AccountType = "Time" Then
            GetDefTimeAccounts = "" & Trim(.rdoColumns(0))
         Else
            GetDefTimeAccounts = "" & Trim(.rdoColumns(1))
         End If
         ClearResultSet RdoAcc
      End With
   End If
   Set RdoAcc = Nothing
   Exit Function
   
modErr1:
   GetDefTimeAccounts = ""
   
End Function


'Public Function GetCardNumber(pempEmployee As Employee, LoginTime As Date) As String
'   'Provides for a unique time card number
'   'Stored as an (11) char string
'   'The First (5) is the card date
'   'the last (6) pickup the time to the part of a
'   'second.
'
'   Dim s As Single
'   Dim l As Long
'   Dim m As Long
'   Dim T As String
'   Dim prdoRes As rdoResultset
'   Dim pstrTmCard As String
'   Dim loginDate As String
'   loginDate = Format(LoginTime, "mm/dd/yyyy")
'
'   ' Check for existing time card for the day, if found use it for all
'   ' other time charges to jobs for the day.
'   sSql = "SELECT TMCARD FROM TchdTable WHERE " _
'          & "TMEMP=" & pempEmployee.intNumber _
'          & " AND TMDAY='" & loginDate & "'"
'   gblnSqlRows = GetDataSet(prdoRes)
'
'   If gblnSqlRows Then
'      With prdoRes
'         pstrTmCard = Trim(.rdoColumns(0))
'      End With
'   Else
'      m = DateValue(Format(GetServerDateTime, "yyyy,mm,dd"))
'      s = TimeValue(Format(GetServerDateTime, "hh:mm:ss"))
'      l = s * 1000000
'      pstrTmCard = Format(m, "00000") & Format(l, "000000")
'      sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDAY) VALUES ('" _
'             & pstrTmCard & "'," & pempEmployee.intNumber & ",'" _
'             & Format(GetServerDateTime, "mm/dd/yyyy") & "')"
'      RdoCon.Execute sSql, rdExecDirect
'   End If
'   Set prdoRes = Nothing
'   GetCardNumber = pstrTmCard
'End Function
'

Public Function GetTimeCardID(EmpNo As Long, startDate As Date)
   'find the existing timecard for employee and date
   'if none exists, create one
   
   Dim dateString As String
   dateString = Format(startDate, "mm/dd/yyyy")
   
   sSql = "select TMCARD from TchdTable" & vbCrLf _
      & "where TMEMP = " & EmpNo & " and TMDAY = '" & dateString & "'"
   Dim rdo As rdoResultset
   If GetDataSet(rdo) Then
      GetTimeCardID = rdo.rdoColumns(0)
      Exit Function
   End If
   
   'timecard doesn't exist.  create it.
   Dim s As Single
   Dim l As Long
   Dim m As Long
   'Dim T As String
   'Dim timeCardNo As String
   Dim loginDate As String
   m = DateValue(Format(GetServerDateTime, "yyyy,mm,dd"))
   s = TimeValue(Format(GetServerDateTime, "hh:mm:ss"))
   l = s * 1000000
   GetTimeCardID = Format(m, "00000") & Format(l, "000000")
   sSql = "INSERT INTO TchdTable (TMCARD,TMEMP,TMDAY)" & vbCrLf _
      & "VALUES ('" & GetTimeCardID & "', " & EmpNo & ", '" & dateString & "')"
   RdoCon.Execute sSql, rdExecDirect
   
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

Public Sub ComputeOverlappingCharges(EmpNo As Long, CardDate As Date)
   
   'if no overlapped time charges, just set elapsed times
   
   Dim rdo As rdoResultset
   Dim sql As String
   Dim where As String
   where = "where TCEMP = " & EmpNo & " and TMDAY = '" & Format(CardDate, "mm/dd/yyyy") & "'" & vbCrLf _
      & "and TCSURUN <> 'I'  AND TCSTOPTIME > TCSTARTTIME"
   
   sql = _
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
   
   Dim segments As Integer, jobCount As Integer
   segments = 0
   jobCount = 0
   
   RdoCon.BeginTrans
   sSql = sql
   If GetDataSet(rdo) Then
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
   sSql = "Update TcitTable set TCHOURS = cast(cast(datediff(n, TCSTARTTIME, TCSTOPTIME) as decimal(10,3))" & vbCrLf _
      & "/ 60.0  as decimal(10,3))," & vbCrLf _
      & "TCMULTIJOB = 0" & vbCrLf _
      & "from TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & where
   RdoCon.Execute sSql
   
   'if no overlapping segments, we're done
   If jobCount <= 1 Then
      RdoCon.CommitTrans
      Exit Sub
   End If
   
   'there are overlapping charges.  subtract shared amounts for each segment
   segments = 0
   jobCount = 0
   Dim prevSegmentStart As Date, segmentStart As Date
   sSql = sql
   If GetDataSet(rdo) Then
      Do While Not rdo.EOF
         segmentStart = rdo!OpTime
         If segments > 1 Then
            Dim Hours As Currency
            Hours = (segments - 1) * DateDiff("n", prevSegmentStart, segmentStart) / (60 * segments)
            If Hours > 0 Then
               sSql = "update TcitTable set TCHOURS = cast(TCHOURS - " & Hours & " as decimal(10,3))," & vbCrLf _
                  & "TCMULTIJOB = case when TCMULTIJOB < " & segments & " then " & segments & " else TCMULTIJOB end" & vbCrLf _
                  & "where TCEMP = " & EmpNo & vbCrLf _
                  & "and TCSTARTTIME <= '" & prevSegmentStart & "'" & vbCrLf _
                  & "and TCSTOPTIME >= '" & segmentStart & "'"
'Debug.Print prevSegmentStart & " - " & segmentStart & " count " & segments & " subtract " & Hours
               RdoCon.Execute sSql
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

   RdoCon.CommitTrans
End Sub

Sub UpdateTimeCardTotals(EmployeeNumber As Long, CardDate As Date)
   'update daily totals for timecard
   'TchdTable.TMREGHRS,TMOVTHRS,TMDBLHRS,TMSTART,TMSTOP

   sSql = "UpdateTimeCardTotals " & EmployeeNumber & ", '" & Format(CardDate, "mm/dd/yyyy") & "'"
   RdoCon.Execute sSql, rdExecDirect
End Sub


Function ApplyShiftCode(iEmpNum As Long, strCardDate As String, _
                        strSftBeg As String, strSftEnd As String, strSfRegHrs As String)
    
    Dim RdoEmp As rdoResultset
    
    On Error Resume Next
    'On Error GoTo DiaErr1
    
    sSql = "SELECT TMCARD, SFSTHR, SFENHR, ISNULL(SFLUNSTHR, 0) SFLUNSTHR, SFADJHR, " & _
            "ISNULL(SFLUNENHR, 0) SFLUNENHR, TMSTART, TMSTOP, TMREGHRS " & _
         "From viewShiftCdEmployeeDetail, TchdTable " & _
        "WHERE TMEMP = PREMNUMBER AND " & _
            "PREMNUMBER = '" & CStr(iEmpNum) & "' AND " & _
            "TMDAY = '" & strCardDate & "'"
    
    Dim minutes As Integer
    Dim strTMCard As String
    Dim strSftBeg, strSftEnd As String
    Dim strSfHrs, strLStart, strAdjHrs As String
    Dim strLBeg, strLEnd, strLunHrs As String
    Dim strBeg, strEnd As String
    Dim strTMBeg, strTMEnd, strTMReg As String
    
    bSqlRows = GetDataSet(RdoEmp, ES_FORWARD)
    If bSqlRows Then
        With RdoEmp
         strTMCard = Trim(!TMCARD)
         strSftBeg = Trim(!SFSTHR)
         strSftEnd = Trim(!SFENHR)
         strAdjHrs = Trim(!SFADJHR)
         strLStart = Trim(!SFLUNSTHR)
         strLEnd = Trim(!SFLUNENHR)
         strTMBeg = Trim(!TMSTART)
         strTMEnd = Trim(!TMSTOP)
         strTMReg = Trim(!TMREGHRS)
         
         ClearResultSet RdoEmp
        End With
        
        Dim where As String
        where = "where TMCARD = '" & strTMCard & "' AND TCEMP = " & CStr(iEmpNum) & vbCrLf _
            & " and TMDAY = '" & Format(strCardDate, "mm/dd/yyyy") & "'" & vbCrLf _
           & "and TCSURUN <> 'I'  AND TCSTOPTIME > TCSTARTTIME"
        
        ' Do we need to adjust
        Dim diff1 As Integer
        Dim dtStart As Date
        diff1 = Abs(DateDiff("n", strSftBeg, strTMBeg))
        
        If CInt(diff1) <= CInt(strAdjHrs) Then
            
            sSql = "UPDATE TchdTable SET TMSTART = '" & strSftBeg & "'" & _
                " WHERE TMCARD = '" & strTMCard & "' AND TMEMP = '" & CStr(iEmpNum) & "'" & _
                    " AND TMDAY = '" & strCardDate & "'"
            RdoCon.Execute sSql, rdExecDirect
            
            dtStart = CDate(strCardDate & " " & strSftBeg)
            sSql = "UPDATE TcitTable SET TCSTART = '" & strSftBeg & "'," & _
                    " TCSTARTTIME = '" & Format(dtStart, "mm/dd/yy hh:mm") & "' " & _
                   " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
                  & where & " AND TCSTART = '" & strTMBeg & "'"
            RdoCon.Execute sSql, rdExecDirect
            
            strBeg = strSftBeg
        Else
            strBeg = strTMBeg
        End If
        
        Dim diff As Integer
        Dim dtEnd As Date
        diff = Abs(DateDiff("n", strSftEnd, strTMEnd))
        If CInt(diff) <= CInt(strAdjHrs) Then
            
            sSql = "UPDATE TchdTable SET TMSTOP = '" & strSftEnd & "'" & _
                " WHERE TMCARD = '" & strTMCard & "' AND TMEMP = '" & CStr(iEmpNum) & "'" & _
                    " AND TMDAY = '" & strCardDate & "'"
            RdoCon.Execute sSql, rdExecDirect
        
            dtEnd = CDate(strCardDate & " " & strSftEnd)
            sSql = "UPDATE TcitTable SET TCSTOP = '" & strSftEnd & "'," & _
                    " TCSTOPTIME = '" & Format(dtEnd, "mm/dd/yy hh:mm") & "' " & _
                   " FROM TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
                  & where & " AND TCSTOP = '" & strTMEnd & "'"
            RdoCon.Execute sSql, rdExecDirect
        
            strEnd = strSftEnd
        Else
            strEnd = strTMEnd
        End If
        
        ' Adjust the Luch time
        AdjustLunchTime iEmpNum, strTMCard, CDate(strCardDate), CStr(strBeg), _
                                        CStr(strEnd), CStr(strLStart), CStr(strLEnd)
        ' Compute overlap time
        ComputeOverlappingChargesWithLunch iEmpNum, CDate(strCardDate), CStr(strBeg), _
                                CStr(strEnd), CStr(strLStart), CStr(strLEnd)
        
    End If
End Function

Public Sub AdjustLunchTime(iEmpNum As Long, strTMCard As String, strCardDate As String, _
                 strBeg As String, strEnd As String, strLStart As String, strLEnd As String)

    ' First adjust the Time Charge header data
    ' Get Lunch Minutes
    Dim rdo As rdoResultset
    Dim lMinutes As Long
    Dim lHrs As Double
    
    lMinutes = GetLunchMinutes(strBeg, strEnd, strLStart, strLEnd)
    If (Abs(lMinutes) > 0) Then
        lHrs = Round((Abs(lMinutes) / 60#), 3)
    Else
        lHrs = Round(0, 3)
    End If

    sSql = "UPDATE TchdTable SET TMREGHRS = TMREGHRS - " & CStr(lHrs) _
            & " WHERE TMCARD = '" & strTMCard & "' AND TMEMP = '" & CStr(iEmpNum) & "'" & _
                " AND TMDAY = '" & strCardDate & "'"
                
    RdoCon.Execute sSql, rdExecDirect
    
    ' Adjust the TC IT records
    
    sSql = "SELECT DISTINCT TCSTART, TCSTOP, TCHOURS " & _
             " FROM TcitTable WHERE TCEMP = '" & CStr(iEmpNum) & "' AND " & _
            " TCCARD = '" & strTMCard & "'"
   
    If GetDataSet(rdo) Then
        While Not rdo.EOF
            
            Dim strTcBeg As String
            Dim strTcEnd As String
            
            strTcBeg = Trim(rdo!TCSTART)
            strTcEnd = Trim(rdo!TCSTOP)
            
            lMinutes = GetLunchMinutes(strTcBeg, strTcEnd, strLStart, strLEnd)
            
            'calculate full segment elapsed times
            sSql = "Update TcitTable set TCHOURS = cast(cast((datediff(n, TCSTARTTIME, TCSTOPTIME) - " & CStr(lMinutes) & ") as decimal(10,3))" & vbCrLf _
               & "/ 60.0  as decimal(10,3))," & vbCrLf _
               & " TCMULTIJOB = 0 " & vbCrLf _
               & " FROM TcitTable WHERE TCCARD = '" & strTMCard & "' AND TCEMP = '" & CStr(iEmpNum) & "'" & vbCrLf _
               & " AND TCSTART = '" & strTcBeg & "' AND TCSTOP = '" & strTcEnd & "'"
            RdoCon.Execute sSql
            
            rdo.MoveNext
        Wend
    End If
    
End Sub
Public Sub ComputeOverlappingChargesWithLunch(EmpNo As Long, CardDate As Date, _
                                strBeg As String, strEnd As String, strLStart As String, strLEnd As String)
   
   'if no overlapped time charges, just set elapsed times
   
   Dim rdo As rdoResultset
   Dim sql As String
   Dim where As String
   where = "where TCEMP = " & EmpNo & " and TMDAY = '" & Format(CardDate, "mm/dd/yyyy") & "'" & vbCrLf _
      & "and TCSURUN <> 'I'  AND TCSTOPTIME > TCSTARTTIME"
   
   sql = _
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
   
   Dim segments As Integer, jobCount As Integer
   segments = 0
   jobCount = 0
   
   'RdoCon.BeginTrans
   sSql = sql
   If GetDataSet(rdo) Then
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
   
   ' Get Lunch Minutes
   Dim lMinutes As Long
   lMinutes = 0 'GetLunchMinutes(strBeg, strEnd, strLStart, strLEnd)
   
   'calculate full segment elapsed times
   sSql = "Update TcitTable set TCHOURS = cast(cast((datediff(n, TCSTARTTIME, TCSTOPTIME) - " & CStr(lMinutes) & ") as decimal(10,3))" & vbCrLf _
      & "/ 60.0  as decimal(10,3))," & vbCrLf _
      & "TCMULTIJOB = 0" & vbCrLf _
      & "from TcitTable join TchdTable on TCCARD = TMCARD" & vbCrLf _
      & where
   ' MM RdoCon.Execute sSql
   
   'if no overlapping segments, we're done
   If jobCount <= 1 Then
      'RdoCon.CommitTrans
      Exit Sub
   End If
   
   'there are overlapping charges.  subtract shared amounts for each segment
   segments = 0
   jobCount = 0
   Dim prevSegmentStart As Date, segmentStart As Date
   sSql = sql
   If GetDataSet(rdo) Then
      Do While Not rdo.EOF
         segmentStart = rdo!OpTime
         If segments > 1 Then
            Dim Hours As Currency
            ' Get the Lunch minutes
            Dim strSegBeg, strSegEnd As String
            strSegBeg = Format(prevSegmentStart, "hh:mma/p")
            strSegEnd = Format(segmentStart, "hh:mma/p")
            
            lMinutes = GetLunchMinutes(CStr(strSegBeg), CStr(strSegEnd), strLStart, strLEnd)
            
            Hours = (segments - 1) * (DateDiff("n", prevSegmentStart, segmentStart) - lMinutes) / (60 * segments)
            If Hours > 0 Then
               sSql = "update TcitTable set TCHOURS = cast(TCHOURS - " & Hours & " as decimal(10,3))," & vbCrLf _
                  & "TCMULTIJOB = case when TCMULTIJOB < " & segments & " then " & segments & " else TCMULTIJOB end" & vbCrLf _
                  & "where TCEMP = " & EmpNo & vbCrLf _
                  & "and TCSTARTTIME <= '" & prevSegmentStart & "'" & vbCrLf _
                  & "and TCSTOPTIME >= '" & segmentStart & "'"
'Debug.Print prevSegmentStart & " - " & segmentStart & " count " & segments & " subtract " & Hours
               RdoCon.Execute sSql
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

   'RdoCon.CommitTrans
End Sub

Function GetLunchMinutes(strBeg As String, strEnd As String, strLStart As String, strLEnd As String) As Long
    
    Dim strBeg1, strEnd1 As String
    Dim lLHours As Long
    
    If (CDate(strLStart) >= CDate(strEnd)) Then
        GetLunchMinutes = 0
        Exit Function
    End If
    If (CDate(strLEnd) <= CDate(strBeg)) Then
        GetLunchMinutes = 0
        Exit Function
    End If
    
    If (CDate(strLStart) >= CDate(strBeg)) Then
        strBeg1 = strLStart
    Else
        strBeg1 = strBeg
    End If
    
    If (CDate(strLEnd) >= CDate(strEnd)) Then
        strEnd1 = strEnd
    Else
        strEnd1 = strLEnd
    End If
    
    
   lLHours = DateDiff("n", strBeg1, strEnd1)
   GetLunchMinutes = lLHours

End Function

Function AdjustShiftStartEndTime(ByRef strBeg As String, ByRef strEnd As String, _
                                    ByVal strSfBeg As String, ByVal strSfEnd As String)
    
    If (CDate(strSfBeg) >= CDate(strBeg)) Then
        strBeg = strSfBeg
    End If
    
    If (CDate(strSfEnd) <= CDate(strEnd)) Then
        strEnd = strSfBeg
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


