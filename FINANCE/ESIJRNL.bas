Attribute VB_Name = "ESIJRNL"
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions
' 3/17/05 cjs Added ModErr
Option Explicit

'************************************************************************************************
' ESIJRNL - ES/2000 Journals
'
' Notes:
'   - Journal Types
'       SJ =   Sales
'       PJ =   Purchases
'       CR =   Cash Receipts
'       CC =   Computer Checks
'       XC =   External Checks
'       MC =   Manual Checks (Dropped)
'       PL =   Payroll Labor
'       PD =   Payroll Disbursements
'       TJ =   Time Journal
'       SC =   Sales Commission
'       GL =   General Ledger
'       IJ =   Inventory
'
' Created: 04/23/02 (nth)
' Revisons:
' 09/29/04 (nth) Added JournalInBalance
'
'************************************************************************************************

' Just like GetNextTrans except it use the Journal ID and
' transaction number to determine the next reference number

Public Function GetNextRef(sJrnlId As String, lTrans As Long) As Long
   Dim rdoTrn As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(DCREF) FROM JritTable WHERE DCHEAD='" _
          & Trim(sJrnlId) & "'AND DCTRAN=" & lTrans
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTrn, ES_FORWARD)
   If bSqlRows Then
      With rdoTrn
         If Not IsNull(.Fields(0)) Then
            
            GetNextRef = (.Fields(0)) + 1
         Else
            GetNextRef = 1
         End If
         .Cancel
      End With
   Else
      GetNextRef = 1
   End If
   Set rdoTrn = Nothing
   Exit Function
   
modErr1:
   On Error Resume Next
   sProcName = "GetNextRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

'*************************************************************************
'   Function: GetNextTransaction
'
'   Notes:
'   Code Sample
'   sJournalId = GetOpenJournal("SJ", Format$(Now, "mm/dd/yy"))
'   If Left(sJournalId, 4) = "None" Then sJournalId = ""
'   If sJournalId <> "" Then iTrans = GetNextTransaction(sJournalId)
'
'   Created:
'   Modified:   1/24/01 (nth)
'*************************************************************************

Public Function GetNextTransaction(sJrnlId As String) As Long
   Dim rdoTrn As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(DCTRAN) FROM JritTable WHERE DCHEAD='" _
          & Trim(sJrnlId) & "'"
bSqlRows = clsADOCon.GetDataSet(sSql, rdoTrn, ES_FORWARD)
   If bSqlRows Then
      With rdoTrn
         If Not IsNull(.Fields(0)) Then
            GetNextTransaction = (.Fields(0)) + 1
         Else
            GetNextTransaction = 1
         End If
         .Cancel
      End With
   Else
      GetNextTransaction = 1
   End If
   Exit Function
modErr1:
   On Error Resume Next
   sProcName = "getnexttrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Public Function GetOpenJournal(sJrType As String, sJrDate As String) As String
   Dim rdoJrn As ADODB.Recordset
   Dim rdoJrn1 As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT COREF,COGLVERIFY FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then Co.GlVerify = rdoJrn!COGLVERIFY
   
   Set rdoJrn = Nothing
   ' I'm taking this out for now
   ' Co.GlVerify means to verify accounts of sub journals
   ' when rolling up into GL
   
   'If Co.GlVerify = 1 Then
   'See if there is one
   sSql = "SELECT MJTYPE,MJSTART,MJEND,MJCLOSED,MJGLJRNL FROM JrhdTable " _
          & "WHERE MJTYPE='" & sJrType & "' AND ('" & Format(sJrDate, "mm/dd/yy") & "' " _
          & "BETWEEN MJSTART AND MJEND) AND MJCLOSED IS NULL"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn1, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn1
         If Not IsNull(.Fields(4)) Then
            GetOpenJournal = .Fields(4)
         Else
            GetOpenJournal = ""
         End If
         .Cancel
      End With
   End If
   Set rdoJrn1 = Nothing
   
   sProcName = "getopenjournal"
   Exit Function
modErr1:
   sProcName = "getopenjour"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function

Public Sub CurrentJournal( _
                          sType As String, _
                          ByVal sDate As String, _
                          sJournal As String)
   
   Dim sMsg As String
   Dim b As Byte
   Dim sDesc As String
   Dim bIndex As Byte
   On Error GoTo modErr1
   sDate = Format(sDate, "mm/dd/yy")
   sJournal = GetOpenJournal(sType, sDate)
   If sJournal = "" Then
      sDesc = JournalType(sType, bIndex)
      sMsg = "No Open " & sDesc & " Journal Found For " _
             & sDate & "." & vbCrLf & "Open New Journal For Period?"
      b = MsgBox(sMsg, ES_YESQUESTION, MdiSect.ActiveForm.Caption)
      If b = vbYes Then
         diaJRf01a.bIndex = bIndex
         diaJRf01a.bRemote = True
         diaJRf01a.cmbTyp = sDesc
         diaJRf01a.Show
      End If
   End If
   Exit Sub
modErr1:
   sProcName = "currentjo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Public Function JournalType(sType As String, Optional bIndex As Byte) As String
   Select Case sType
      Case "SJ"
         JournalType = "Sales Journal"
         bIndex = 0
      Case "PJ"
         JournalType = "Purchases Journal"
         bIndex = 1
      Case "CR"
         JournalType = "Cash Receipts Journal"
         bIndex = 2
      Case "CC"
         JournalType = "Computer Checks"
         bIndex = 3
      Case "XC"
         JournalType = "External Checks"
         bIndex = 4
'      Case "PL"
'         JournalType = "Payroll Labor Journal"
'         bIndex = 5
'      Case "PD"
'         JournalType = "Payroll Disbursements Journal"
'         bIndex = 6
      Case "TJ"
         JournalType = "Time Journal"
         bIndex = 5
      Case "IJ"
         JournalType = "Inventory Journal"
         bIndex = 6
   End Select
End Function

Public Function JournalInBalance(sJournal As String) As Byte
   Dim RdoBal As ADODB.Recordset
   sProcName = "journalin"
   On Error GoTo modErr1
   sSql = "SELECT SUM(DCCREDIT),SUM(DCDEBIT) FROM JritTable " _
          & "WHERE DCHEAD = '" & sJournal & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBal)
   If bSqlRows Then
      With RdoBal
         If Format(.Fields(0) & "0", CURRENCYMASK) _
                   = Format(.Fields(1) & "0", CURRENCYMASK) Then
            JournalInBalance = True
         End If
      End With
   End If
   Set RdoBal = Nothing
   Exit Function
modErr1:
   sProcName = "journalinbal"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function
