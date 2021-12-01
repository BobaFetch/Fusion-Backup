Attribute VB_Name = "EsiAcct"
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Customer permissions 6/28/03
'8/28/06 Fixed Account Description
Option Explicit
'Accounting Options added 10/2/00
'If Co.GlVerify = 1 then see if there is a journal open
'in for this date

Public sJournalID As String

Public Function GetThisAccount(sAccount As String) As String
   Dim RdoGlm As ADODB.Recordset
   
   On Error GoTo modErr1
   sAccount = Compress(sAccount)
   sSql = "SELECT GLACCTREF,GLACCTNO FROM GlacTable " _
          & "WHERE GLACCTREF='" & sAccount & "' AND GLINACTIVE=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_FORWARD)
   If bSqlRows Then
      With RdoGlm
         GetThisAccount = Trim(!GLACCTNO)
         ClearResultSet RdoGlm
      End With
   Else
      GetThisAccount = ""
   End If
   Set RdoGlm = Nothing
   Exit Function
   
modErr1:
   Set RdoGlm = Nothing
   Resume modErr2
modErr2:
   Set RdoGlm = Nothing
   On Error GoTo 0
   
End Function

'WIP Labor Accounts 8/28/99
'Use local errors

Public Function GetLaborAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoLac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVLABACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                    Trim(RdoLac.Fields(0))
   End If
   
   'Product Code
   If GetLaborAcct = "" Then
      sSql = "SELECT PCINVLABACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                       Trim(RdoLac.Fields(0))
      End If
   End If
   
   'Default
   If GetLaborAcct = "" Then
      sSql = "SELECT COINVLABACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoLac.Fields(0)) Then GetLaborAcct = _
                       Trim(RdoLac.Fields(0))
      End If
   End If
   Set RdoLac = Nothing
   
End Function

'WIP Overhead Accounts 8/28/99
'Use local errors

Public Function GetOverHeadAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoOac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVOHDACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                    Trim(RdoOac.Fields(0))
   End If
   
   'Product Code
   If GetOverHeadAcct = "" Then
      sSql = "SELECT PCINVOHDACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                       Trim(RdoOac.Fields(0))
      End If
   End If
   
   'Default
   If GetOverHeadAcct = "" Then
      sSql = "SELECT COINVOHDACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoOac.Fields(0)) Then GetOverHeadAcct = _
                       Trim(RdoOac.Fields(0))
      End If
   End If
   Set RdoOac = Nothing
   
End Function

'WIP Material Accounts 8/28/99
'Use local errors

Public Function GetMaterialAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoMac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVMATACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                    Trim(RdoMac.Fields(0))
   End If
   
   'Product Code
   If GetMaterialAcct = "" Then
      sSql = "SELECT PCINVMATACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                       Trim(RdoMac.Fields(0))
      End If
   End If
   
   'Default
   If GetMaterialAcct = "" Then
      sSql = "SELECT COINVMATACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoMac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoMac.Fields(0)) Then GetMaterialAcct = _
                       Trim(RdoMac.Fields(0))
      End If
   End If
   Set RdoMac = Nothing
   
End Function

'WIP Expense Accounts 8/28/99
'Use local errors

Public Function GetExpenseAcct(sPartNumber As String, sCode As String, bLevel As Byte) As String
   Dim RdoEac As ADODB.Recordset
   
   'Part Number
   sSql = "SELECT PAINVEXPACCT FROM PartTable WHERE " _
          & "PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
   If bSqlRows Then
      If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                    Trim(RdoEac.Fields(0))
   End If
   
   'Product Code
   If GetExpenseAcct = "" Then
      sSql = "SELECT PCINVEXPACCT FROM PcodTable WHERE " _
             & "PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                       Trim(RdoEac.Fields(0))
      End If
   End If
   
   'Default
   If GetExpenseAcct = "" Then
      sSql = "SELECT COINVEXPACCT" & Trim(str(bLevel)) & " FROM ComnTable " _
             & "WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEac, ES_FORWARD)
      If bSqlRows Then
         If Not IsNull(RdoEac.Fields(0)) Then GetExpenseAcct = _
                       Trim(RdoEac.Fields(0))
      End If
   End If
   Set RdoEac = Nothing
   Exit Function
   
modErr1:
   Set RdoEac = Nothing
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
   
End Function

'Code Sample
'sJournalId = GetOpenJournal("SJ", format$(es_sysdate, "mm/dd/yy"))
'If Left(sJournalId, 4) = "None" Then sJournalId = ""
'If sJournalId <> "" Then iTrans = GetNextTransaction(sJournalId)

Public Function GetNextTransaction(sJrnlId As String) As Integer
   Dim RdoTrn As ADODB.Recordset
   sSql = "SELECT MAX(DCTRAN) FROM JritTable WHERE DCHEAD='" _
          & Trim(sJrnlId) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrn, ES_FORWARD)
   If bSqlRows Then
      With RdoTrn
         If Not IsNull(.Fields(0)) Then
            GetNextTransaction = (.Fields(0)) + 1
         Else
            GetNextTransaction = 1
         End If
         ClearResultSet RdoTrn
      End With
   Else
      GetNextTransaction = 1
   End If
   Set RdoTrn = Nothing
   Exit Function
   
DiaErr1:
   On Error Resume Next
   Set RdoTrn = Nothing
   sProcName = "getnexttrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'DoModuleErrors MDISect.ActiveForm
   Dim frm As New ClassErrorForm
   DoModuleErrors frm
   
End Function


'Changed to add trap 10/22/03

Public Sub FindAccount(frm As Form)
   Dim RdoAct As ADODB.Recordset
   Dim sAccount As String
   
   On Error GoTo modErr1
   sSql = "Qry_GetAccount '" & Compress(frm.cmbAct) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
   If bSqlRows Then
      With RdoAct
         On Error Resume Next
         frm.cmbAct = "" & Trim(!GLACCTNO)
         frm.lblActdsc = "" & Trim(!GLDESCR)
         If Err > 0 Then _
            frm.lblDsc = "" & Trim(!GLDESCR)
         ClearResultSet RdoAct
      End With
   Else
      If Len(sAccount) > 0 Then frm.lblActdsc = "*** Account Wasn't Found ***"
   End If
   Set RdoAct = Nothing
   Exit Sub
modErr1:
   Set RdoAct = Nothing
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub


'Public Function GetOpenJournal(sJrType As String, sJrDate As String) As String
'   'sJournalId = GetOpenJournal("CR", format$(es_sysdate, "mm/dd/yy"))
'   'See CodeDoc
'   Dim RdoJrn As ADODB.Recordset
'   sSql = "SELECT COREF,COGLVERIFY FROM ComnTable WHERE COREF=1"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoJrn, ES_FORWARD)
'   If bSqlRows Then Co.GlVerify = RdoJrn!COGLVERIFY
'   If Co.GlVerify = 1 Then
'      'See if there is one
'      sSql = "SELECT MJTYPE,MJSTART,MJEND,MJCLOSED,MJGLJRNL FROM JrhdTable " _
'             & "WHERE MJTYPE='" & sJrType & "' AND ('" & sJrDate & "' " _
'             & "BETWEEN MJSTART AND MJEND) AND MJCLOSED IS NULL"
'      bSqlRows = clsADOCon.GetDataSet(sSql, RdoJrn, ES_FORWARD)
'      If bSqlRows Then
'         With RdoJrn
'            If Not IsNull(.Fields(4)) Then
'               GetOpenJournal = .Fields(4)
'            Else
'               GetOpenJournal = ""
'            End If
'            ClearResultSet RdoJrn
'         End With
'      End If
'   Else
'      GetOpenJournal = "None Required"
'   End If
'   sProcName = "getopenjournal"
'
'End Function
'

Public Function GetOpenJournal(sJrType As String, sJrDate As String) As String
   'sJournalId = GetOpenJournal("CR", format$(es_sysdate, "mm/dd/yy"))
   'sJournalId = GetOpenJournal("CR", es_sysdate)
   'See CodeDoc
   
   Dim rdoJrn As ADODB.Recordset
   sSql = "SELECT COREF,COGLVERIFY FROM ComnTable WHERE COREF=1"
   If clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD) Then
      Co.GlVerify = rdoJrn!COGLVERIFY
End If
   
   'See if there is one
   sSql = "select dbo.fnGetOpenJournalID('" & sJrType & "', '" & sJrDate & "')"
   If clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD) Then
      GetOpenJournal = "" & rdoJrn.Fields(0)
   End If
   
   If GetOpenJournal = "" And Co.GlVerify = 0 Then
      GetOpenJournal = "None Reqd"
   End If
   
   Set rdoJrn = Nothing
End Function


Public Sub CodeDoc()
   '   sJournalId = GetOpenJournal("CR", format$(es_sysdate, "mm/dd/yy"))
   '        If Left(sJournalId, 4) = "None" Then
   '            sJournalId = ""
   '            b = 1
   '        Else
   '            If sJournalId = "" Then b = 0 Else b = 1
   '        End If
   '        If b = 0 Then
   '            MsgBox "There Is No Open Cash Receipts Journal For The Period.", _
   '                vbExclamation, Caption
   '            Sleep 500
   '            Unload Me
   '            Exit Sub
   '        Else
   
End Sub


'Grab accounts for invoices
'local errors

Public Function GetPartInvoiceAccounts(sPartNumber As String, iLevel As Integer, sCode As String, sREVAccount As String, sCGSAccount As String) As Boolean
   Dim RdoAct As ADODB.Recordset
   'Parts
   GetPartInvoiceAccounts = True
   sPartNumber = Compress(sPartNumber)
   sSql = "SELECT PACGSMATACCT,PAREVACCT FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
   If bSqlRows Then
      With RdoAct
         sCGSAccount = "" & Trim(!PACGSMATACCT)
         sREVAccount = "" & Trim(!PAREVACCT)
         ClearResultSet RdoAct
      End With
   End If
   
   If sCGSAccount = "" Or sREVAccount = "" Then
      'product code
      sCode = Compress(sCode)
      sSql = "SELECT PCCGSMATACCT,PCREVACCT FROM PcodTable " _
             & "WHERE PCREF='" & sCode & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
      If bSqlRows Then
         With RdoAct
            If sCGSAccount = "" Then sCGSAccount = "" & Trim(!PCCGSMATACCT)
            If sREVAccount = "" Then sREVAccount = "" & Trim(!PCREVACCT)
            ClearResultSet RdoAct
         End With
      End If
   End If
   
   If sCGSAccount = "" Or sREVAccount = "" Then
      'Company
      sSql = "SELECT COCGSACCT" & Trim(str(iLevel)) & "," _
             & "COREVACCT" & Trim(str(iLevel)) & " FROM " _
             & "ComnTable WHERE COREF=1"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAct, ES_FORWARD)
      If bSqlRows Then
         With RdoAct
            If sCGSAccount = "" Then sCGSAccount = "" & Trim(.Fields(0))
            If sREVAccount = "" Then sREVAccount = "" & Trim(.Fields(1))
            ClearResultSet RdoAct
         End With
      End If
   End If
   If sCGSAccount = "" Or sREVAccount = "" Then GetPartInvoiceAccounts = False
   Set RdoAct = Nothing
   Exit Function
   
modErr1:
   Set RdoAct = Nothing
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Function

Public Sub Documentation()
   'SJ = SALES JOURNAL
   'PJ = PURCHASES JOURNAL
   'CR = CASH RECEIPTS JOURNAL
   'CC = CASH DISBURSEMENTS JNL - COMPUTER CHECKS
   'XC = CASH DISBURSEMENTS JNL - EXTERNAL CHECKS
   'PL = PAYROLL LABOR JOURNAL
   'PD = PAYROLL DISBURSEMENTS JOURNAL
   'TJ = TIME JOURNAL
   'IJ = INVENTORY JOURNAL
   'GL - GENERAL LEDGER JOURNALS
   'CT - CASH TRANSFER JOURNALS
   'OF - AR/AP OFFSET JOURNALS
   
End Sub
