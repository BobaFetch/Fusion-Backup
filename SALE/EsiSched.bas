Attribute VB_Name = "EsiSched"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'Sales/Production (MO Creation)
'12/22/05 Added to Allow MO's to be created at Sales Order Items (SaleSLe02b)
'
Option Explicit
Public bGoodSoMo As Byte
Public iAutoIncr As Integer
Public vTimeFormat As Variant

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


Public Sub FillRoutings()
   Dim RdoRtg As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "Qry_FillRoutings "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRtg, ES_FORWARD)
   If bSqlRows Then
      With RdoRtg
         MdiSect.ActiveForm.cmbRte = "" & Trim(!RTNUM)
         Do Until .EOF
            AddComboStr MdiSect.ActiveForm.cmbRte.hwnd, "" & Trim(!RTNUM)
            .MoveNext
         Loop
         ClearResultSet RdoRtg
      End With
   End If
   Set RdoRtg = Nothing
   Exit Sub
   
modErr1:
   sProcName = "fillroutings"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

'See if there are calendars
'Syntax is bGoodCal = GetCenterCalendar(Me, Format$(SomeDate,"mm/dd/yy")
'6/1/00

Public Function GetCenterCalendar(frm As Form, Optional sMonth As String) As Boolean
   On Error Resume Next
   If sMonth = "" Then
      sMonth = Format(ES_SYSDATE, "mmm") & "-" & Format(ES_SYSDATE, "yyyy")
   Else
      sMonth = Format(sMonth, "mmm") & "-" & Format(sMonth, "yyyy")
   End If
   sSql = "SELECT WCCREF FROM WcclTable WHERE " _
          & "WCCREF='" & sMonth & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   If clsADOCon.RowsAffected = 0 Then
      GetCenterCalendar = False
      MsgBox "There Are No Work Center Calendars " & vbCrLf _
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
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   Set RdoTme = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)

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
   
   'Blocked 8/17/00 per Gary(pct)/Colin
   '        If GetCenterCalHours > 0 Then
   '            cResources = GetCenterCalRes(sMonth, sShop, sCenter, iDay)
   '            GetCenterCalHours = cResources * GetCenterCalHours
   '        End If
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
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   Set RdoTme = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)
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
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
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
   'Set RdoTme = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
   Set RdoTme = clsADOCon.GetRecordSet(sSql, adOpenForwardOnly)

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

Public Function GetCompanyCalendar() As Byte
   Dim RdoCal As ADODB.Recordset
   Dim sCalYear As String
   Dim sCalMonth As String
   
   On Error Resume Next
   sCalYear = Format$(Now, "yyyy")
   sCalMonth = Format$(Now, "mmm")
   sSql = "SELECT COCREF FROM CoclTable WHERE COCREF='" _
          & sCalMonth & "-" & sCalYear & " '"
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
